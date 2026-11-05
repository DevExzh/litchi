//! The formatting and structure operations of `formatting` (L-DOCX-W5):
//! per-property splices with preservation, exact no-ops and inverses,
//! structure round trips, refusals, composition, durable replay, the
//! bounded-growth layout differential, and the DOCX fixture corpus.

use super::*;
use crate::color::Theme;
use crate::format::UnderlineStyle;
use crate::paragraph::{LineSpacingRule, ParagraphSpacing, RunUnderline, RunUnderlineColor};
use crate::properties::{
    Border, BorderStyle, Color, EighthPoints, Emphasis, FontHint, Fonts, HalfPoints, HexColor,
    Highlight, Indentation, Justification, Language, Numbering, ParagraphBorders, Points, Shading,
    ShadingPattern, TabKind, TabLeader, TabStop, TextAlignment, TextDirection, TextEffect,
    TextboxTightWrap, ThemeFont, Twips, VerticalAlign,
};
use litchi_core::patch::{BlobLimits, CompositionLimits, PatchLimits};

const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WORD_STRICT: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const W14: &str = "http://schemas.microsoft.com/office/word/2010/wordml";

fn document(body: &str) -> Vec<u8> {
    format!(
        r#"<w:document xmlns:w="{WORD}" xmlns:w14="{W14}" xmlns:x="urn:extension"><w:body>{body}</w:body></w:document>"#
    )
    .into_bytes()
}

fn snapshot(body: &str) -> Snapshot {
    Snapshot::from_xml(document(body)).unwrap()
}

fn limits() -> PatchLimits {
    PatchLimits::new(
        BlobLimits::new(64, 32 * 1024 * 1024, 64 * 1024 * 1024),
        64 * 1024 * 1024,
        4_096,
        16,
        16 * 1024 * 1024,
        64 * 1024 * 1024,
    )
}

/// The retained layout of a snapshot equals a rescan of its bytes.
fn assert_layout_matches_rescan(snapshot: &Snapshot) {
    let layout = scan_document(snapshot.xml_bytes()).unwrap();
    let same = |left: &[Range], right: &[Range]| {
        left.len() == right.len()
            && left
                .iter()
                .zip(right)
                .all(|(left, right)| left.start == right.start && left.length == right.length)
    };
    assert!(same(&snapshot.paragraphs, &layout.paragraphs));
    assert!(same(&snapshot.tables, &layout.tables));
    assert!(same(&snapshot.block_controls, &layout.block_controls));
    assert_eq!(snapshot.content_end, layout.content_end);
    assert_eq!(
        snapshot.paragraph_namespaces.len(),
        layout.paragraph_namespaces.len()
    );
    for (retained, scanned) in snapshot
        .paragraph_namespaces
        .iter()
        .zip(&layout.paragraph_namespaces)
    {
        assert_eq!(retained.as_ref(), scanned.as_ref());
    }
}

/// Commit, then prove the patch is exactly reversible three ways: the
/// in-memory inverse, the durable forward replay and the durable inverse
/// replay (both directions through deterministic JSON).
fn commit_and_check(base: &Snapshot, edit: Edit) -> Commit {
    assert_layout_matches_rescan(edit.projected());
    let commit = edit.commit().unwrap();
    assert_layout_matches_rescan(commit.snapshot());
    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.xml_bytes(), base.xml_bytes());
    if !commit.patch().is_noop() {
        let durable = commit.patch().to_durable(limits()).unwrap();
        let json = durable.to_deterministic_json().unwrap();
        let durable =
            litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
                &json,
                limits(),
            )
            .unwrap();
        let replayed = base.apply_durable(&durable).unwrap();
        assert_eq!(replayed.xml_bytes(), commit.snapshot().xml_bytes());
        let inverse = commit.patch().inverse().to_durable(limits()).unwrap();
        let undone = commit.snapshot().apply_durable(&inverse).unwrap();
        assert_eq!(undone.xml_bytes(), base.xml_bytes());
    }
    commit
}

fn paragraph_xml(snapshot: &Snapshot, position: usize) -> Vec<u8> {
    snapshot
        .paragraph(Position::new(position))
        .unwrap()
        .xml_bytes()
        .to_vec()
}

fn run_xml(snapshot: &Snapshot, paragraph: usize, run: usize) -> Vec<u8> {
    let paragraph = paragraph_xml(snapshot, paragraph);
    let (start, end) = formatting::direct_run_for_tests(&paragraph, run);
    paragraph[start..end].to_vec()
}

fn run_properties(snapshot: &Snapshot, paragraph: usize, run: usize) -> crate::properties::Run {
    crate::properties::parse_run(
        &run_xml(snapshot, paragraph, run),
        &snapshot.paragraph_namespaces[paragraph],
    )
    .unwrap()
}

fn paragraph_properties(snapshot: &Snapshot, paragraph: usize) -> crate::properties::Paragraph {
    crate::properties::parse_paragraph(
        &paragraph_xml(snapshot, paragraph),
        &snapshot.paragraph_namespaces[paragraph],
    )
    .unwrap()
}

/// The direct child elements of the first `container` element of `xml`, as
/// written, in order.
fn container_children(xml: &[u8], container: &str) -> Vec<Vec<u8>> {
    let open = format!("<w:{container}");
    let Some(start) = find(xml, open.as_bytes()) else {
        return Vec::new();
    };
    let mut reader = Reader::from_reader(&xml[start..]);
    let mut depth = 0usize;
    let mut children = Vec::new();
    let mut child_start = 0usize;
    loop {
        let before = usize::try_from(reader.buffer_position()).unwrap();
        let event = reader.read_event().unwrap();
        let after = usize::try_from(reader.buffer_position()).unwrap();
        match event {
            Event::Start(_) => {
                depth += 1;
                if depth == 2 {
                    child_start = before;
                }
            },
            Event::Empty(_) => {
                if depth == 0 {
                    return children;
                }
                if depth == 1 {
                    children.push(xml[start + before..start + after].to_vec());
                }
            },
            Event::End(_) => {
                if depth == 2 {
                    children.push(xml[start + child_start..start + after].to_vec());
                }
                depth -= 1;
                if depth == 0 {
                    return children;
                }
            },
            Event::Eof => return children,
            _ => {},
        }
    }
}

fn find(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
}

fn local_name(element: &[u8]) -> String {
    let end = element
        .iter()
        .position(|byte| matches!(byte, b' ' | b'/' | b'>'))
        .unwrap();
    let name = std::str::from_utf8(&element[1..end]).unwrap();
    name.rsplit(':').next().unwrap().to_owned()
}

/// Property preservation concerns the root paragraph/run's direct property
/// container. A drawing may contain nested text-box paragraphs with their
/// own properties; those belong to the unchanged-content assertion instead.
fn direct_property_children(xml: &[u8], container: &str) -> Vec<Vec<u8>> {
    let mut reader = Reader::from_reader(xml);
    let mut depth = 0usize;
    let mut selected = false;
    let mut child_start = 0usize;
    let mut children = Vec::new();
    loop {
        let before = usize::try_from(reader.buffer_position()).unwrap();
        let event = reader.read_event().unwrap();
        let after = usize::try_from(reader.buffer_position()).unwrap();
        match event {
            Event::Start(element) => {
                if depth == 1 && element.local_name().as_ref() == container.as_bytes() {
                    selected = true;
                }
                depth += 1;
                if selected && depth == 3 {
                    child_start = before;
                }
            },
            Event::Empty(element) => {
                if selected && depth == 2 {
                    children.push(xml[before..after].to_vec());
                } else if depth == 1 && element.local_name().as_ref() == container.as_bytes() {
                    return children;
                }
            },
            Event::End(_) => {
                if selected && depth == 3 {
                    children.push(xml[child_start..after].to_vec());
                } else if selected && depth == 2 {
                    return children;
                }
                depth -= 1;
            },
            Event::Eof => return children,
            _ => {},
        }
    }
}

#[test]
fn property_preservation_oracle_excludes_nested_text_box_properties() {
    let nested = "<w:r><w:pict><w:txbxContent><w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr></w:p></w:txbxContent></w:pict></w:r>";
    let before = format!("<w:p>{nested}</w:p>");
    let after =
        format!("<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/><w:keepNext/></w:pPr>{nested}</w:p>");
    assert!(direct_property_children(before.as_bytes(), "pPr").is_empty());
    assert_eq!(
        direct_property_children(after.as_bytes(), "pPr"),
        [
            b"<w:pStyle w:val=\"Heading1\"/>".to_vec(),
            b"<w:keepNext/>".to_vec()
        ]
    );
    assert_preserved(
        before.as_bytes(),
        after.as_bytes(),
        "pPr",
        &["pStyle", "keepNext"],
    );
}

/// Every child of the container that is not `changed` survives byte for
/// byte and in order.
fn assert_preserved(before: &[u8], after: &[u8], container: &str, changed: &[&str]) {
    let keep = |children: Vec<Vec<u8>>| {
        children
            .into_iter()
            .filter(|child| !changed.contains(&local_name(child).as_str()))
            .collect::<Vec<_>>()
    };
    assert_eq!(
        keep(direct_property_children(before, container)),
        keep(direct_property_children(after, container)),
        "unchanged children of w:{container} moved or changed"
    );
}

fn refusal(result: Result<impl Sized>) -> Refusal {
    match result {
        Err(Error::Refused { reason, .. }) => reason,
        Err(error) => panic!("expected a refusal, got {error:?}"),
        Ok(_) => panic!("expected a refusal"),
    }
}

// ---------------------------------------------------------------------------
// The typed property vocabulary.
// ---------------------------------------------------------------------------

fn color(rgb: [u8; 3]) -> Color {
    Color {
        value: HexColor::Rgb(rgb),
        theme: None,
        theme_tint: None,
        theme_shade: None,
    }
}

fn border(style: &str, size: i64, rgb: [u8; 3]) -> Border {
    let mut border = Border::new(BorderStyle::from_xml(style).unwrap());
    border.size = Some(EighthPoints(size));
    border.color = Some(HexColor::Rgb(rgb));
    border.space = Some(Points(0));
    border
}

fn shading(pattern: ShadingPattern, fill: [u8; 3]) -> Shading {
    let mut shading = Shading::new(pattern);
    shading.fill = Some(HexColor::Rgb(fill));
    shading
}

type RunSetter = (
    &'static str,
    &'static [&'static str],
    fn(&mut crate::properties::Run),
);

/// One setter per typed run property, with the element(s) it changes.
fn run_setters() -> Vec<RunSetter> {
    vec![
        ("rStyle", &["rStyle"], |run| {
            run.style = Some("Emphasis".into())
        }),
        ("rFonts", &["rFonts"], |run| {
            run.fonts = Some(Fonts {
                hint: Some(FontHint::EastAsia),
                ascii: Some("Georgia".into()),
                h_ansi: Some("Georgia".into()),
                east_asia: Some("SimSun".into()),
                complex_script: Some("Arial".into()),
                ascii_theme: Some(ThemeFont::MinorHAnsi),
                h_ansi_theme: Some(ThemeFont::MinorHAnsi),
                east_asia_theme: Some(ThemeFont::MinorEastAsia),
                complex_script_theme: Some(ThemeFont::MinorBidi),
            });
        }),
        ("b", &["b"], |run| run.bold = Some(true)),
        ("bCs", &["bCs"], |run| run.bold_complex_script = Some(true)),
        ("i false", &["i"], |run| run.italic = Some(false)),
        ("iCs", &["iCs"], |run| {
            run.italic_complex_script = Some(true)
        }),
        ("caps", &["caps"], |run| run.caps = Some(true)),
        ("smallCaps", &["smallCaps"], |run| {
            run.small_caps = Some(true)
        }),
        ("strike", &["strike"], |run| run.strike = Some(true)),
        ("dstrike", &["dstrike"], |run| {
            run.double_strike = Some(true)
        }),
        ("outline", &["outline"], |run| run.outline = Some(true)),
        ("shadow", &["shadow"], |run| run.shadow = Some(true)),
        ("emboss", &["emboss"], |run| run.emboss = Some(true)),
        ("imprint", &["imprint"], |run| run.imprint = Some(true)),
        ("noProof", &["noProof"], |run| run.no_proof = Some(true)),
        ("snapToGrid", &["snapToGrid"], |run| {
            run.snap_to_grid = Some(false);
        }),
        ("vanish", &["vanish"], |run| run.vanish = Some(true)),
        ("webHidden", &["webHidden"], |run| {
            run.web_hidden = Some(true)
        }),
        ("color", &["color"], |run| {
            run.color = Some(Color {
                value: HexColor::Rgb([0x1F, 0x4E, 0x79]),
                theme: Some(Theme::Accent1),
                theme_tint: Some(0x99),
                theme_shade: Some(0xBF),
            });
        }),
        ("color auto", &["color"], |run| {
            run.color = Some(Color {
                value: HexColor::Auto,
                theme: None,
                theme_tint: None,
                theme_shade: None,
            });
        }),
        ("spacing", &["spacing"], |run| {
            run.character_spacing = Some(Twips(-40));
        }),
        ("w", &["w"], |run| run.scale = Some(150)),
        ("kern", &["kern"], |run| run.kerning = Some(HalfPoints(28))),
        ("position", &["position"], |run| {
            run.position = Some(HalfPoints(-6));
        }),
        ("sz", &["sz"], |run| run.size = Some(HalfPoints(28))),
        ("szCs", &["szCs"], |run| {
            run.size_complex_script = Some(HalfPoints(28));
        }),
        ("highlight", &["highlight"], |run| {
            run.highlight = Some(Highlight::Yellow);
        }),
        ("u", &["u"], |run| {
            run.underline = Some(RunUnderline {
                style: UnderlineStyle::Wave,
                color: Some(RunUnderlineColor::Rgb([0xFF, 0, 0])),
                theme_color: Some(Theme::Accent2),
                theme_tint: Some(0x10),
                theme_shade: Some(0x20),
            });
        }),
        ("effect", &["effect"], |run| {
            run.effect = Some(TextEffect::BlinkBackground);
        }),
        ("bdr", &["bdr"], |run| {
            run.border = Some(border("single", 4, [0xFF, 0, 0]));
        }),
        ("shd", &["shd"], |run| {
            run.shading = Some(shading(ShadingPattern::Clear, [0xFF, 0xFF, 0]));
        }),
        ("vertAlign", &["vertAlign"], |run| {
            run.vertical_align = Some(VerticalAlign::Superscript);
        }),
        ("rtl", &["rtl"], |run| run.right_to_left = Some(true)),
        ("cs", &["cs"], |run| run.complex_script = Some(true)),
        ("em", &["em"], |run| run.emphasis = Some(Emphasis::Dot)),
        ("lang", &["lang"], |run| {
            run.language = Some(Language {
                latin: Some("en-US".into()),
                east_asia: Some("zh-CN".into()),
                bidi: Some("ar-SA".into()),
            });
        }),
        ("specVanish", &["specVanish"], |run| {
            run.special_vanish = Some(true);
        }),
        ("oMath", &["oMath"], |run| run.office_math = Some(true)),
    ]
}

type ParagraphSetter = (
    &'static str,
    &'static [&'static str],
    fn(&mut crate::properties::Paragraph),
);

/// One setter per typed paragraph property, with the element(s) it changes.
fn paragraph_setters() -> Vec<ParagraphSetter> {
    vec![
        ("pStyle", &["pStyle"], |paragraph| {
            paragraph.style = Some("Heading2".into());
        }),
        ("keepNext", &["keepNext"], |paragraph| {
            paragraph.keep_next = Some(true);
        }),
        ("keepLines", &["keepLines"], |paragraph| {
            paragraph.keep_lines = Some(true);
        }),
        ("pageBreakBefore", &["pageBreakBefore"], |paragraph| {
            paragraph.page_break_before = Some(true);
        }),
        ("widowControl", &["widowControl"], |paragraph| {
            paragraph.widow_control = Some(false);
        }),
        ("numPr", &["numPr"], |paragraph| {
            paragraph.numbering = Some(Numbering {
                id: Some(4),
                level: Some(2),
            });
        }),
        (
            "suppressLineNumbers",
            &["suppressLineNumbers"],
            |paragraph| {
                paragraph.suppress_line_numbers = Some(true);
            },
        ),
        ("pBdr", &["pBdr"], |paragraph| {
            paragraph.borders = Some(ParagraphBorders {
                top: Some(border("single", 8, [0, 0x70, 0xC0])),
                bottom: Some(border("double", 4, [0, 0, 0])),
                between: Some(border("dotted", 2, [1, 2, 3])),
                ..ParagraphBorders::default()
            });
        }),
        ("shd", &["shd"], |paragraph| {
            let mut value = shading(ShadingPattern::Percent15, [0xDD, 0xEB, 0xF7]);
            value.color = Some(HexColor::Rgb([0xC0, 0, 0]));
            value.theme_fill = Some(Theme::Accent3);
            paragraph.shading = Some(value);
        }),
        ("tabs", &["tabs"], |paragraph| {
            paragraph.tabs = Some(vec![
                TabStop::new(TabKind::Left, Twips(720)),
                TabStop::new(TabKind::Decimal, Twips(1440)).with_leader(TabLeader::Dot),
            ]);
        }),
        (
            "suppressAutoHyphens",
            &["suppressAutoHyphens"],
            |paragraph| {
                paragraph.suppress_auto_hyphens = Some(true);
            },
        ),
        ("kinsoku", &["kinsoku"], |paragraph| {
            paragraph.kinsoku = Some(false);
        }),
        ("wordWrap", &["wordWrap"], |paragraph| {
            paragraph.word_wrap = Some(false);
        }),
        ("overflowPunct", &["overflowPunct"], |paragraph| {
            paragraph.overflow_punctuation = Some(true);
        }),
        ("topLinePunct", &["topLinePunct"], |paragraph| {
            paragraph.top_line_punctuation = Some(true);
        }),
        ("autoSpaceDE", &["autoSpaceDE"], |paragraph| {
            paragraph.auto_space_east_asian_latin = Some(false);
        }),
        ("autoSpaceDN", &["autoSpaceDN"], |paragraph| {
            paragraph.auto_space_east_asian_numbers = Some(false);
        }),
        ("bidi", &["bidi"], |paragraph| paragraph.bidi = Some(true)),
        ("adjustRightInd", &["adjustRightInd"], |paragraph| {
            paragraph.adjust_right_indent = Some(false);
        }),
        ("snapToGrid", &["snapToGrid"], |paragraph| {
            paragraph.snap_to_grid = Some(false);
        }),
        ("spacing", &["spacing"], |paragraph| {
            paragraph.spacing = Some(ParagraphSpacing {
                before: Some(360),
                before_lines: Some(100),
                before_auto_spacing: Some(false),
                after: Some(120),
                after_lines: None,
                after_auto_spacing: Some(true),
                line: Some(360),
                line_rule: Some(LineSpacingRule::Auto),
            });
        }),
        ("ind", &["ind"], |paragraph| {
            paragraph.indentation = Some(Indentation {
                left: Some(Twips(567)),
                right: Some(Twips(1134)),
                hanging: Some(Twips(567)),
                first_line_chars: Some(200),
                ..Indentation::default()
            });
        }),
        ("contextualSpacing", &["contextualSpacing"], |paragraph| {
            paragraph.contextual_spacing = Some(true);
        }),
        ("mirrorIndents", &["mirrorIndents"], |paragraph| {
            paragraph.mirror_indents = Some(true);
        }),
        ("suppressOverlap", &["suppressOverlap"], |paragraph| {
            paragraph.suppress_overlap = Some(true);
        }),
        ("jc", &["jc"], |paragraph| {
            paragraph.justification = Some(Justification::Both);
        }),
        ("textDirection", &["textDirection"], |paragraph| {
            paragraph.text_direction = Some(TextDirection::TbRl);
        }),
        ("textAlignment", &["textAlignment"], |paragraph| {
            paragraph.text_alignment = Some(TextAlignment::Center);
        }),
        ("textboxTightWrap", &["textboxTightWrap"], |paragraph| {
            paragraph.textbox_tight_wrap = Some(TextboxTightWrap::AllLines);
        }),
        ("outlineLvl", &["outlineLvl"], |paragraph| {
            paragraph.outline_level = Some(1);
        }),
        ("mark", &["rPr"], |paragraph| {
            let mark = paragraph.mark.get_or_insert_with(Default::default);
            mark.bold = Some(true);
            mark.color = Some(color([0xC0, 0, 0]));
            mark.size = Some(HalfPoints(28));
        }),
    ]
}

// ---------------------------------------------------------------------------
// Run properties.
// ---------------------------------------------------------------------------

/// A run whose `w:rPr` holds an untyped WordprocessingML property, a Word
/// 2010 extension element and a tracked formatting change.
const NOISY_RUN: &str = concat!(
    r#"<w:p><w:r><w:rPr><w:rStyle w:val="Strong"/><w:i w:val="true"/>"#,
    r#"<w:eastAsianLayout w:id="7" w:vert="1"/>"#,
    r#"<w14:glow w14:rad="63500"><w14:schemeClr w14:val="accent1"/></w14:glow>"#,
    r#"<w:rPrChange w:id="9" w:author="Ann" w:date="2026-01-01T00:00:00Z"><w:rPr><w:b/></w:rPr></w:rPrChange>"#,
    r#"</w:rPr><w:t>text</w:t></w:r></w:p><w:sectPr/>"#
);

#[test]
fn every_typed_run_property_splices_reads_back_and_preserves_its_neighbours() {
    let base = snapshot(NOISY_RUN);
    let before_run = run_xml(&base, 0, 0);
    for (name, changed, set) in run_setters() {
        let mut expected = run_properties(&base, 0, 0);
        set(&mut expected);
        let mut edit = base.edit();
        edit.format_run(Position::new(0), Position::new(0), set)
            .unwrap_or_else(|error| panic!("{name}: {error}"));
        let commit = commit_and_check(&base, edit);
        let after = commit.snapshot();
        assert!(
            crate::properties::same_run(&run_properties(after, 0, 0), &expected),
            "{name}: typed readback differs"
        );
        let after_run = run_xml(after, 0, 0);
        assert_preserved(&before_run, &after_run, "rPr", changed);
        let children = container_children(&after_run, "rPr");
        assert_eq!(
            local_name(children.last().unwrap()),
            "rPrChange",
            "{name}: the tracked change must stay last"
        );
        let glow = children
            .iter()
            .position(|child| local_name(child) == "glow")
            .unwrap();
        let written = children
            .iter()
            .position(|child| local_name(child) == changed[0])
            .unwrap();
        assert!(
            written < glow,
            "{name}: a new WordprocessingML property goes before the extension element"
        );
        assert_eq!(
            Paragraph::new(paragraph_xml(after, 0)).text().unwrap(),
            "text"
        );

        // Removing the property again removes exactly its element.
        let mut edit = after.edit();
        let current = run_properties(after, 0, 0);
        edit.format_run(Position::new(0), Position::new(0), |run| {
            let mut cleared = current.clone();
            set(&mut cleared);
            *run = cleared;
            clear_run_property(run, name);
        })
        .unwrap();
        let removed = commit_and_check(after, edit);
        if !["rStyle", "i false"].contains(&name) {
            assert_eq!(
                removed.snapshot().xml_bytes(),
                base.xml_bytes(),
                "{name}: set then remove must restore the bytes"
            );
        }
    }
}

fn clear_run_property(run: &mut crate::properties::Run, name: &str) {
    match name {
        "rStyle" => run.style = None,
        "rFonts" => run.fonts = None,
        "b" => run.bold = None,
        "bCs" => run.bold_complex_script = None,
        "i false" => run.italic = None,
        "iCs" => run.italic_complex_script = None,
        "caps" => run.caps = None,
        "smallCaps" => run.small_caps = None,
        "strike" => run.strike = None,
        "dstrike" => run.double_strike = None,
        "outline" => run.outline = None,
        "shadow" => run.shadow = None,
        "emboss" => run.emboss = None,
        "imprint" => run.imprint = None,
        "noProof" => run.no_proof = None,
        "snapToGrid" => run.snap_to_grid = None,
        "vanish" => run.vanish = None,
        "webHidden" => run.web_hidden = None,
        "color" | "color auto" => run.color = None,
        "spacing" => run.character_spacing = None,
        "w" => run.scale = None,
        "kern" => run.kerning = None,
        "position" => run.position = None,
        "sz" => run.size = None,
        "szCs" => run.size_complex_script = None,
        "highlight" => run.highlight = None,
        "u" => run.underline = None,
        "effect" => run.effect = None,
        "bdr" => run.border = None,
        "shd" => run.shading = None,
        "vertAlign" => run.vertical_align = None,
        "rtl" => run.right_to_left = None,
        "cs" => run.complex_script = None,
        "em" => run.emphasis = None,
        "lang" => run.language = None,
        "specVanish" => run.special_vanish = None,
        "oMath" => run.office_math = None,
        other => panic!("no clearer for {other}"),
    }
}

#[test]
fn a_run_without_properties_gains_a_container_first() {
    let base = snapshot(r#"<w:p><w:r><w:t>a</w:t></w:r><w:r/></w:p>"#);
    let mut edit = base.edit();
    edit.format_run(Position::new(0), Position::new(0), |run| {
        run.bold = Some(true);
        run.size = Some(HalfPoints(24));
    })
    .unwrap();
    edit.format_run(Position::new(0), Position::new(1), |run| {
        run.italic = Some(true);
    })
    .unwrap();
    let commit = commit_and_check(&base, edit);
    assert_eq!(
        paragraph_xml(commit.snapshot(), 0),
        br#"<w:p><w:r><w:rPr><w:b/><w:sz w:val="24"/></w:rPr><w:t>a</w:t></w:r><w:r><w:rPr><w:i/></w:rPr></w:r></w:p>"#
    );
}

#[test]
fn unchanged_values_are_exact_noops_and_keep_their_lexical_form() {
    let base = snapshot(
        r#"<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="12pt"/></w:pPr><w:r><w:rPr><w:b w:val="true"/><w:sz w:val="12pt"/><w:color w:val="1f4e79"/></w:rPr><w:t>x</w:t></w:r></w:p>"#,
    );
    let mut edit = base.edit();
    edit.format_run(Position::new(0), Position::new(0), |run| {
        run.bold = Some(true);
        run.size = Some(HalfPoints(24));
        run.color = Some(color([0x1F, 0x4E, 0x79]));
    })
    .unwrap();
    edit.format_run(Position::new(0), Position::new(0), |_run| {})
        .unwrap();
    edit.format_paragraph(Position::new(0), |paragraph| {
        paragraph.justification = Some(Justification::Center);
        paragraph.spacing.as_mut().unwrap().before = Some(240);
    })
    .unwrap();
    assert!(edit.operations.is_empty());
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.snapshot().xml_bytes(), base.xml_bytes());
}

#[test]
fn a_changed_compound_property_rewrites_only_its_changed_attributes() {
    let base = snapshot(
        r#"<w:p><w:r><w:rPr><w:rFonts  w:hAnsi='Arial' x:extra="kept" w:cs="Times New Roman" w:ascii="Arial" w:eastAsia="SimSun"/></w:rPr><w:t>x</w:t></w:r></w:p>"#,
    );
    let mut edit = base.edit();
    edit.format_run(Position::new(0), Position::new(0), |run| {
        let fonts = run.fonts.as_mut().unwrap();
        fonts.ascii = Some("Calibri".into());
        fonts.h_ansi = Some("Calibri".into());
        fonts.east_asia_theme = Some(ThemeFont::MinorEastAsia);
    })
    .unwrap();
    let commit = commit_and_check(&base, edit);
    assert_eq!(
        run_xml(commit.snapshot(), 0, 0),
        br#"<w:r><w:rPr><w:rFonts  w:hAnsi="Calibri" x:extra="kept" w:cs="Times New Roman" w:ascii="Calibri" w:eastAsia="SimSun" w:eastAsiaTheme="minorEastAsia"/></w:rPr><w:t>x</w:t></w:r>"#
    );
}

#[test]
fn toggles_switch_lexically_minimal_and_an_invalid_property_survives_until_set() {
    let base = snapshot(
        r#"<w:p><w:r><w:rPr><w:b w:val="0"/><w:i/><w:sz w:val="large"/></w:rPr><w:t>x</w:t></w:r></w:p>"#,
    );
    assert_eq!(run_properties(&base, 0, 0).size, None);
    let mut edit = base.edit();
    edit.format_run(Position::new(0), Position::new(0), |run| {
        run.bold = Some(true);
        run.italic = Some(false);
    })
    .unwrap();
    let toggled = commit_and_check(&base, edit);
    assert_eq!(
        run_xml(toggled.snapshot(), 0, 0),
        br#"<w:r><w:rPr><w:b/><w:i w:val="0"/><w:sz w:val="large"/></w:rPr><w:t>x</w:t></w:r>"#
    );
    let mut edit = toggled.snapshot().edit();
    edit.format_run(Position::new(0), Position::new(0), |run| {
        run.size = Some(HalfPoints(20));
    })
    .unwrap();
    let sized = commit_and_check(toggled.snapshot(), edit);
    assert_eq!(
        run_xml(sized.snapshot(), 0, 0),
        br#"<w:r><w:rPr><w:b/><w:i w:val="0"/><w:sz w:val="20"/></w:rPr><w:t>x</w:t></w:r>"#
    );
}

// ---------------------------------------------------------------------------
// Paragraph properties.
// ---------------------------------------------------------------------------

/// A paragraph whose `w:pPr` holds untyped WordprocessingML properties, a
/// tracked paragraph-mark insertion with a tracked mark formatting change,
/// and a tracked paragraph property change.
const NOISY_PARAGRAPH: &str = concat!(
    r#"<w:p><w:pPr><w:pStyle w:val="Normal"/>"#,
    r#"<w:framePr w:w="4320" w:h="720" w:wrap="around"/>"#,
    r#"<w:divId w:val="123456"/><w:cnfStyle w:val="000000000000"/>"#,
    r#"<w:rPr><w:ins w:id="3" w:author="Ann" w:date="2026-01-01T00:00:00Z"/><w:i/>"#,
    r#"<w:rPrChange w:id="4" w:author="Ann" w:date="2026-01-01T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr>"#,
    r#"<w:pPrChange w:id="5" w:author="Ann" w:date="2026-01-01T00:00:00Z"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange>"#,
    r#"</w:pPr><w:r><w:t>body</w:t></w:r></w:p><w:sectPr/>"#
);

#[test]
fn every_typed_paragraph_property_splices_reads_back_and_preserves_its_neighbours() {
    let base = snapshot(NOISY_PARAGRAPH);
    let before_paragraph = paragraph_xml(&base, 0);
    for (name, changed, set) in paragraph_setters() {
        let mut expected = paragraph_properties(&base, 0);
        set(&mut expected);
        let mut edit = base.edit();
        edit.format_paragraph(Position::new(0), set)
            .unwrap_or_else(|error| panic!("{name}: {error}"));
        let commit = commit_and_check(&base, edit);
        let after = commit.snapshot();
        assert!(
            crate::properties::same_paragraph(&paragraph_properties(after, 0), &expected),
            "{name}: typed readback differs"
        );
        let after_paragraph = paragraph_xml(after, 0);
        assert_preserved(&before_paragraph, &after_paragraph, "pPr", changed);
        let children = container_children(&after_paragraph, "pPr");
        assert_eq!(local_name(children.last().unwrap()), "pPrChange", "{name}");
        // The paragraph mark keeps its tracked insertion first and its
        // tracked formatting change last.
        let mark = container_children(&after_paragraph, "rPr");
        assert_eq!(local_name(&mark[0]), "ins", "{name}");
        assert_eq!(local_name(mark.last().unwrap()), "rPrChange", "{name}");
        // The schema order of the known children holds.
        let ranks = children
            .iter()
            .map(|child| crate::properties::paragraph_rank_for_tests(&local_name(child)).unwrap())
            .collect::<Vec<_>>();
        assert!(ranks.windows(2).all(|pair| pair[0] < pair[1]), "{name}");

        // Removing the property again removes exactly what was written.
        let mut edit = after.edit();
        let original = paragraph_properties(&base, 0);
        edit.format_paragraph(Position::new(0), |paragraph| {
            clear_paragraph_property(paragraph, name, &original);
        })
        .unwrap();
        let removed = commit_and_check(after, edit);
        assert_eq!(
            removed.snapshot().xml_bytes(),
            base.xml_bytes(),
            "{name}: set then remove must restore the bytes"
        );
    }
}

/// Undo one paragraph setter: a property the base paragraph states goes back
/// to its value, any other is removed.
fn clear_paragraph_property(
    paragraph: &mut crate::properties::Paragraph,
    name: &str,
    original: &crate::properties::Paragraph,
) {
    match name {
        "pStyle" => paragraph.style.clone_from(&original.style),
        "keepNext" => paragraph.keep_next = None,
        "keepLines" => paragraph.keep_lines = None,
        "pageBreakBefore" => paragraph.page_break_before = None,
        "widowControl" => paragraph.widow_control = None,
        "numPr" => paragraph.numbering = None,
        "suppressLineNumbers" => paragraph.suppress_line_numbers = None,
        "pBdr" => paragraph.borders = None,
        "shd" => paragraph.shading = None,
        "tabs" => paragraph.tabs = None,
        "suppressAutoHyphens" => paragraph.suppress_auto_hyphens = None,
        "kinsoku" => paragraph.kinsoku = None,
        "wordWrap" => paragraph.word_wrap = None,
        "overflowPunct" => paragraph.overflow_punctuation = None,
        "topLinePunct" => paragraph.top_line_punctuation = None,
        "autoSpaceDE" => paragraph.auto_space_east_asian_latin = None,
        "autoSpaceDN" => paragraph.auto_space_east_asian_numbers = None,
        "bidi" => paragraph.bidi = None,
        "adjustRightInd" => paragraph.adjust_right_indent = None,
        "snapToGrid" => paragraph.snap_to_grid = None,
        "spacing" => paragraph.spacing = None,
        "ind" => paragraph.indentation = None,
        "contextualSpacing" => paragraph.contextual_spacing = None,
        "mirrorIndents" => paragraph.mirror_indents = None,
        "suppressOverlap" => paragraph.suppress_overlap = None,
        "jc" => paragraph.justification = None,
        "textDirection" => paragraph.text_direction = None,
        "textAlignment" => paragraph.text_alignment = None,
        "textboxTightWrap" => paragraph.textbox_tight_wrap = None,
        "outlineLvl" => paragraph.outline_level = None,
        "mark" => paragraph.mark.clone_from(&original.mark),
        other => panic!("no clearer for {other}"),
    }
}

#[test]
fn style_and_numbering_application_round_trip() {
    let base = snapshot(
        r#"<w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/><w:ins w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"/></w:numPr></w:pPr><w:r><w:t>second</w:t></w:r></w:p>"#,
    );
    let mut edit = base.edit();
    edit.format_paragraph(Position::new(0), |paragraph| {
        paragraph.style = Some("Heading1".into());
        paragraph.numbering = Some(Numbering {
            id: Some(2),
            level: Some(0),
        });
    })
    .unwrap();
    edit.format_run(Position::new(0), Position::new(0), |run| {
        run.style = Some("Emphasis".into());
    })
    .unwrap();
    edit.format_paragraph(Position::new(1), |paragraph| {
        paragraph.numbering.as_mut().unwrap().id = Some(7);
        paragraph.numbering.as_mut().unwrap().level = Some(3);
    })
    .unwrap();
    let commit = commit_and_check(&base, edit);
    let first = Paragraph::new(paragraph_xml(commit.snapshot(), 0));
    assert_eq!(first.style_id().unwrap().as_deref(), Some("Heading1"));
    let numbering = first.numbering().unwrap().unwrap();
    assert_eq!((numbering.num_id, numbering.level), (2, 0));
    assert_eq!(
        run_properties(commit.snapshot(), 0, 0).style.as_deref(),
        Some("Emphasis")
    );
    assert_eq!(
        paragraph_xml(commit.snapshot(), 1),
        br#"<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="3"/><w:numId w:val="7"/><w:ins w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"/></w:numPr></w:pPr><w:r><w:t>second</w:t></w:r></w:p>"#
    );
    // Removing a numbering that carries a tracked numbering insertion would
    // drop the revision.
    let mut edit = commit.snapshot().edit();
    assert_eq!(
        refusal(edit.format_paragraph(Position::new(1), |paragraph| {
            paragraph.numbering = None;
        })),
        Refusal::TrackedChange
    );
    assert!(edit.operations.is_empty());
}

#[test]
fn property_containers_are_removed_only_when_nothing_untyped_goes_with_them() {
    let base = snapshot(concat!(
        r#"<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/><x:tab/></w:tabs>"#,
        r#"<w:rPr><w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"/><w:b/></w:rPr></w:pPr>"#,
        r#"<w:r><w:t>x</w:t></w:r></w:p>"#,
        r#"<w:p><w:pPr><w:pBdr><w:top w:val="single"/></w:pBdr><w:rPr><w:b/></w:rPr></w:pPr></w:p>"#
    ));
    let mut edit = base.edit();
    assert_eq!(
        refusal(edit.format_paragraph(Position::new(0), |paragraph| paragraph.tabs = None)),
        Refusal::ComplexContent
    );
    assert_eq!(
        refusal(edit.format_paragraph(Position::new(0), |paragraph| paragraph.mark = None)),
        Refusal::TrackedChange
    );
    // The mark's own properties still change around its tracked deletion.
    edit.format_paragraph(Position::new(0), |paragraph| {
        paragraph.mark.as_mut().unwrap().bold = None;
        paragraph.mark.as_mut().unwrap().italic = Some(true);
    })
    .unwrap();
    edit.format_paragraph(Position::new(1), |paragraph| {
        paragraph.borders = None;
        paragraph.mark = None;
    })
    .unwrap();
    let commit = commit_and_check(&base, edit);
    assert_eq!(
        container_children(&paragraph_xml(commit.snapshot(), 0), "rPr"),
        [
            br#"<w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"/>"#.to_vec(),
            b"<w:i/>".to_vec()
        ]
    );
    assert_eq!(
        paragraph_xml(commit.snapshot(), 1),
        b"<w:p><w:pPr></w:pPr></w:p>"
    );
}

#[test]
fn duplicated_property_elements_are_refused_only_when_they_must_change() {
    let base = snapshot(
        r#"<w:p><w:pPr><w:jc w:val="left"/><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:b/><w:b w:val="0"/></w:rPr><w:t>x</w:t></w:r></w:p>"#,
    );
    let mut edit = base.edit();
    assert_eq!(
        refusal(edit.format_run(Position::new(0), Position::new(0), |run| {
            run.bold = Some(false);
        })),
        Refusal::DuplicateProperty
    );
    assert_eq!(
        refusal(edit.format_paragraph(Position::new(0), |paragraph| {
            paragraph.justification = None;
        })),
        Refusal::DuplicateProperty
    );
    edit.format_run(Position::new(0), Position::new(0), |run| {
        run.italic = Some(true);
    })
    .unwrap();
    commit_and_check(&base, edit);
}

#[test]
fn values_outside_the_readers_domain_are_refused() {
    let base = snapshot(r#"<w:p><w:r><w:t>x</w:t></w:r></w:p>"#);
    let mut edit = base.edit();
    let cases: [fn(&mut crate::properties::Run); 4] = [
        |run| run.size = Some(HalfPoints(-1)),
        |run| run.scale = Some(601),
        |run| run.style = Some("bad\u{1}".into()),
        |run| run.border = Some(border("single", -8, [0, 0, 0])),
    ];
    for case in cases {
        assert_eq!(
            refusal(edit.format_run(Position::new(0), Position::new(0), case)),
            Refusal::InvalidPropertyValue
        );
    }
    assert_eq!(
        refusal(edit.format_paragraph(Position::new(0), |paragraph| {
            paragraph.outline_level = Some(10);
        })),
        Refusal::InvalidPropertyValue
    );
    assert_eq!(
        refusal(edit.format_paragraph(Position::new(0), |paragraph| {
            paragraph.spacing = Some(ParagraphSpacing {
                before: Some(u64::MAX),
                ..ParagraphSpacing::default()
            });
        })),
        Refusal::InvalidPropertyValue
    );
    assert!(edit.operations.is_empty());
    assert_eq!(edit.projected().xml_bytes(), base.xml_bytes());
}

#[test]
fn a_strict_document_refuses_transitional_spellings() {
    let base = Snapshot::from_xml(format!(
        r#"<w:document xmlns:w="{WORD_STRICT}"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>"#
    ))
    .unwrap();
    let mut edit = base.edit();
    let cases: [fn(&mut crate::properties::Paragraph); 4] = [
        |paragraph| paragraph.justification = Some(Justification::Left),
        |paragraph| {
            paragraph.indentation = Some(Indentation {
                left: Some(Twips(10)),
                ..Indentation::default()
            });
        },
        |paragraph| paragraph.tabs = Some(vec![TabStop::new(TabKind::Right, Twips(1))]),
        |paragraph| paragraph.text_direction = Some(TextDirection::LrTb),
    ];
    for case in cases {
        assert_eq!(
            refusal(edit.format_paragraph(Position::new(0), case)),
            Refusal::TransitionalValue
        );
    }
    edit.format_paragraph(Position::new(0), |paragraph| {
        paragraph.justification = Some(Justification::Start);
        paragraph.indentation = Some(Indentation {
            start: Some(Twips(720)),
            ..Indentation::default()
        });
    })
    .unwrap();
    let commit = commit_and_check(&base, edit);
    assert_eq!(
        paragraph_xml(commit.snapshot(), 0),
        br#"<w:p><w:pPr><w:ind w:start="720"/><w:jc w:val="start"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#
    );
}

#[test]
fn an_unprefixed_owner_is_refused_rather_than_written_without_namespace() {
    let base = Snapshot::from_xml(format!(
        r#"<document xmlns="{WORD}"><body><p><r><t>x</t></r></p></body></document>"#
    ))
    .unwrap();
    let mut edit = base.edit();
    assert_eq!(
        refusal(edit.format_run(Position::new(0), Position::new(0), |run| {
            run.size = Some(HalfPoints(20));
        })),
        Refusal::NamespaceBinding
    );
    // An element without attributes needs no attribute prefix.
    edit.format_run(Position::new(0), Position::new(0), |run| {
        run.bold = Some(true)
    })
    .unwrap();
    // New content declares its own prefix.
    edit.append_run(Position::new(0), &NewRun::new("y"))
        .unwrap();
    edit.insert_formatted_paragraph(
        BodyAnchor::End,
        &NewParagraph::new().with_run(NewRun::new("z")),
    )
    .unwrap();
    let commit = commit_and_check(&base, edit);
    let xml = std::str::from_utf8(commit.snapshot().xml_bytes()).unwrap();
    assert!(xml.contains(&format!(
        r#"<p><r><rPr><b/></rPr><t>x</t></r><w:r xmlns:w="{WORD}"><w:t>y</w:t></w:r></p>"#
    )));
    assert!(xml.contains(&format!(
        r#"<w:p xmlns:w="{WORD}"><w:r><w:t>z</w:t></w:r></w:p>"#
    )));
}

#[test]
fn budget_managed_snapshots_refuse_the_formatting_operations() {
    use litchi_core::{Budget, CancellationSource, ExecutionLimits, Limits, OwnedSource};
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};
    use litchi_opc::{BlobPart, OpcPackage, PackageWriter};

    let mut opc = OpcPackage::new();
    opc.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        document(r#"<w:p><w:r><w:t>x</w:t></w:r></w:p>"#),
    )))
    .unwrap();
    opc.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    let memory = 1 << 20;
    let budget = Budget::root(
        "docx-managed-formatting",
        Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (_cancellation_source, cancellation) = CancellationSource::pair();
    let execution = ExecutionLimits::new(
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroU64::new(memory).unwrap(),
        0,
    )
    .unwrap();
    let package = crate::source_backed::Package::from_read_at_with_options(
        Arc::new(OwnedSource::new(PackageWriter::to_bytes(&opc).unwrap())),
        crate::OpenOptions::default().execution_context(ExecutionContext::new(
            budget,
            cancellation,
            execution,
        )),
    )
    .unwrap();
    let mut edit = package.edit().unwrap();
    assert!(edit.projected().unwrap().is_managed());
    assert!(matches!(
        edit.format_run(Position::new(0), Position::new(0), |run| run.bold =
            Some(true)),
        Err(Error::Document(crate::Error::Refused(_)))
    ));
    assert!(matches!(
        edit.insert_formatted_paragraph(BodyAnchor::End, &NewParagraph::new()),
        Err(Error::Document(crate::Error::Refused(_)))
    ));
}

// ---------------------------------------------------------------------------
// Structure.
// ---------------------------------------------------------------------------

const BLOCKS: &str = concat!(
    r#"<w:p><w:r><w:t>p0</w:t></w:r></w:p>"#,
    r#"<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    r#"<w:bookmarkEnd w:id="99"/>"#,
    r#"<w:p><w:r><w:t>p1</w:t></w:r></w:p>"#,
    r#"<w:sdt><w:sdtContent><w:p><w:r><w:t>control</w:t></w:r></w:p></w:sdtContent></w:sdt>"#,
    r#"<w:p><w:r><w:t>p2</w:t></w:r></w:p>"#,
    r#"<w:sectPr><w:pgSz w:w="12240"/></w:sectPr>"#
);

fn texts(snapshot: &Snapshot) -> Vec<String> {
    snapshot
        .paragraphs()
        .iter()
        .map(|paragraph| paragraph.text().unwrap())
        .collect()
}

#[test]
fn formatted_paragraphs_insert_at_every_kind_of_anchor() {
    let base = snapshot(BLOCKS);
    let mut new = NewParagraph::new()
        .with_run(NewRun::new("plain "))
        .with_run({
            let mut run = NewRun::new("bold\ttab");
            run.properties_mut().bold = Some(true);
            run
        });
    new.properties_mut().style = Some("Heading1".into());
    new.properties_mut().mark = Some(Box::default());
    let cases = [
        (
            BodyAnchor::Before(BodyBlock::Paragraph(Position::new(0))),
            0,
            "<w:body><w:p><w:pPr>",
        ),
        (
            BodyAnchor::After(BodyBlock::Paragraph(Position::new(0))),
            1,
            "</w:p><w:p><w:pPr><w:pStyle",
        ),
        (
            BodyAnchor::Before(BodyBlock::Table(Position::new(0))),
            1,
            "</w:p><w:p><w:pPr><w:pStyle",
        ),
        (
            BodyAnchor::After(BodyBlock::Table(Position::new(0))),
            1,
            "</w:tbl><w:p><w:pPr>",
        ),
        (
            BodyAnchor::Before(BodyBlock::Paragraph(Position::new(1))),
            1,
            r#"<w:bookmarkEnd w:id="99"/><w:p><w:pPr>"#,
        ),
        (
            BodyAnchor::After(BodyBlock::ContentControl(Position::new(0))),
            2,
            "</w:sdt><w:p><w:pPr>",
        ),
        (BodyAnchor::End, 3, "</w:p><w:p><w:pPr><w:pStyle"),
    ];
    for (anchor, position, marker) in cases {
        let mut edit = base.edit();
        let inserted = edit.insert_formatted_paragraph(anchor, &new).unwrap();
        assert_eq!(inserted, Position::new(position), "{anchor:?}");
        let commit = commit_and_check(&base, edit);
        let after = commit.snapshot();
        assert_eq!(after.paragraph_count(), 4);
        assert_eq!(texts(after)[position], "plain bold\ttab");
        let xml = std::str::from_utf8(after.xml_bytes()).unwrap();
        assert!(xml.contains(marker), "{anchor:?}: {xml}");
        assert!(
            xml.ends_with(r#"<w:sectPr><w:pgSz w:w="12240"/></w:sectPr></w:body></w:document>"#)
        );
        let properties = paragraph_properties(after, position);
        assert_eq!(properties.style.as_deref(), Some("Heading1"));
        assert!(properties.mark.is_some());
        assert_eq!(run_properties(after, position, 1).bold, Some(true));
    }
    let mut edit = base.edit();
    assert_eq!(
        refusal(edit.insert_formatted_paragraph(
            BodyAnchor::Before(BodyBlock::Table(Position::new(5))),
            &new
        )),
        Refusal::BlockNotFound
    );
}

#[test]
fn paragraphs_insert_into_bodies_without_section_properties_or_paragraphs() {
    for body in ["", "<w:tbl/>", "<w:p/>"] {
        let base = snapshot(body);
        let mut edit = base.edit();
        let position = edit
            .insert_formatted_paragraph(BodyAnchor::End, &NewParagraph::new())
            .unwrap();
        let commit = commit_and_check(&base, edit);
        assert_eq!(
            commit.snapshot().paragraph(position).unwrap().xml_bytes(),
            b"<w:p/>"
        );
        assert!(
            commit
                .snapshot()
                .xml_bytes()
                .ends_with(b"<w:p/></w:body></w:document>")
        );
    }
}

#[test]
fn removing_paragraphs_round_trips_and_refuses_what_it_would_break() {
    let base = snapshot(concat!(
        r#"<w:p><w:r><w:t>plain</w:t></w:r></w:p>"#,
        r#"<w:p><w:bookmarkStart w:id="1" w:name="a"/><w:r><w:t>closed bookmark</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>"#,
        r#"<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>section</w:t></w:r></w:p>"#,
        r#"<w:p><w:bookmarkStart w:id="2" w:name="b"/><w:r><w:t>open</w:t></w:r></w:p>"#,
        r#"<w:p><w:r><w:t>end of b</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>"#,
        r#"<w:p><w:r><w:commentReference w:id="0"/></w:r></w:p>"#,
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r></w:p>"#,
        r#"<w:p><w:pPr><w:rPr><w:del w:id="5" w:author="A" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr></w:p>"#,
        r#"<w:p><w:moveTo w:id="6" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:t>moved</w:t></w:r></w:moveTo></w:p>"#,
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
        r#"<w:sectPr/>"#
    ));
    for (position, expected) in [
        (2, Refusal::SectionBreak),
        (3, Refusal::OpenRange),
        (4, Refusal::OpenRange),
        (5, Refusal::StoryReference),
        (6, Refusal::OpenRange),
        (7, Refusal::TrackedChange),
        (8, Refusal::StoryReference),
    ] {
        let mut edit = base.edit();
        assert_eq!(
            refusal(edit.remove_paragraph(Position::new(position))),
            expected,
            "paragraph {position}"
        );
        assert_eq!(edit.projected().xml_bytes(), base.xml_bytes());
    }
    let mut edit = base.edit();
    edit.remove_paragraph(Position::new(9)).unwrap();
    edit.remove_paragraph(Position::new(1)).unwrap();
    edit.remove_paragraph(Position::new(0)).unwrap();
    let commit = commit_and_check(&base, edit);
    assert_eq!(commit.snapshot().paragraph_count(), 7);
    assert_eq!(texts(commit.snapshot())[0], "section");
    // A complete document can be emptied.
    let base = snapshot(r#"<w:p/><w:sectPr/>"#);
    let mut edit = base.edit();
    edit.remove_paragraph(Position::new(0)).unwrap();
    let commit = commit_and_check(&base, edit);
    assert_eq!(commit.snapshot().paragraph_count(), 0);
}

#[test]
fn moving_paragraphs_round_trips_through_every_anchor() {
    let base = snapshot(BLOCKS);
    let cases = [
        (0, BodyAnchor::End, 2, vec!["p1", "p2", "p0"]),
        (
            2,
            BodyAnchor::Before(BodyBlock::Paragraph(Position::new(0))),
            0,
            vec!["p2", "p0", "p1"],
        ),
        (
            0,
            BodyAnchor::After(BodyBlock::Table(Position::new(0))),
            0,
            vec!["p0", "p1", "p2"],
        ),
        (
            2,
            BodyAnchor::Before(BodyBlock::Table(Position::new(0))),
            1,
            vec!["p0", "p2", "p1"],
        ),
        (
            1,
            BodyAnchor::After(BodyBlock::Paragraph(Position::new(2))),
            2,
            vec!["p0", "p2", "p1"],
        ),
        (
            1,
            BodyAnchor::Before(BodyBlock::ContentControl(Position::new(0))),
            1,
            vec!["p0", "p1", "p2"],
        ),
    ];
    for (from, to, moved_to, order) in cases {
        let mut edit = base.edit();
        let moved = edit.move_paragraph(Position::new(from), to).unwrap();
        assert_eq!(moved, Position::new(moved_to), "{from} -> {to:?}");
        let commit = commit_and_check(&base, edit);
        assert_eq!(texts(commit.snapshot()), order, "{from} -> {to:?}");
    }
    // Anchoring a paragraph where it already is records nothing.
    for (from, to) in [
        (
            1,
            BodyAnchor::Before(BodyBlock::Paragraph(Position::new(1))),
        ),
        (1, BodyAnchor::After(BodyBlock::Paragraph(Position::new(1)))),
        (2, BodyAnchor::End),
        (
            2,
            BodyAnchor::After(BodyBlock::ContentControl(Position::new(0))),
        ),
    ] {
        let mut edit = base.edit();
        assert_eq!(
            edit.move_paragraph(Position::new(from), to).unwrap(),
            Position::new(from)
        );
        assert!(edit.operations.is_empty());
    }
    // References travel with a moved paragraph; open ranges do not.
    let base = snapshot(concat!(
        r#"<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>"#,
        r#"<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>x</w:t></w:r></w:p>"#,
        r#"<w:p><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>"#,
    ));
    let mut edit = base.edit();
    edit.move_paragraph(Position::new(0), BodyAnchor::End)
        .unwrap();
    assert_eq!(
        refusal(edit.move_paragraph(Position::new(0), BodyAnchor::End)),
        Refusal::OpenRange
    );
    commit_and_check(&base, edit);
}

#[test]
fn runs_insert_append_and_remove_with_exact_inverses() {
    let base = snapshot(concat!(
        r#"<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>E = mc</w:t></w:r><w:bookmarkEnd w:id="3"/></w:p>"#,
        r#"<w:p/>"#,
        r#"<w:p><w:r><w:t>a</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>b</w:t></w:r></w:p>"#
    ));
    let mut superscript = NewRun::new("2");
    superscript.properties_mut().vertical_align = Some(VerticalAlign::Superscript);
    let mut edit = base.edit();
    assert_eq!(
        edit.append_run(Position::new(0), &superscript).unwrap(),
        Position::new(1)
    );
    assert_eq!(
        edit.insert_run(Position::new(0), Position::new(0), &NewRun::new("> "))
            .unwrap(),
        Position::new(0)
    );
    assert_eq!(
        edit.append_run(Position::new(1), &NewRun::new("into empty"))
            .unwrap(),
        Position::new(0)
    );
    edit.remove_run(Position::new(2), Position::new(1)).unwrap();
    assert_eq!(
        refusal(edit.insert_run(Position::new(2), Position::new(5), &NewRun::new("x"))),
        Refusal::RunNotFound
    );
    let commit = commit_and_check(&base, edit);
    let after = commit.snapshot();
    assert_eq!(
        paragraph_xml(after, 0),
        br#"<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">&gt; </w:t></w:r><w:r><w:t>E = mc</w:t></w:r><w:bookmarkEnd w:id="3"/><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>2</w:t></w:r></w:p>"#
    );
    assert_eq!(
        paragraph_xml(after, 1),
        br#"<w:p><w:r><w:t>into empty</w:t></w:r></w:p>"#
    );
    assert_eq!(
        paragraph_xml(after, 2),
        br#"<w:p><w:r><w:t>a</w:t></w:r></w:p>"#
    );
}

#[test]
fn runs_and_paragraphs_move_losslessly_and_replay_exactly() {
    let base = snapshot(concat!(
        r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>A</w:t><w:opaque w:value="keep"/></w:r>"#,
        "\n<!-- run-owned -->\n",
        r#"<w:r><w:t>B</w:t></w:r></w:p>"#,
        r#"<w:p><w:r><w:t>C</w:t></w:r><w:opaque w:value="paragraph"/></w:p>"#,
        r#"<w:p><w:r><w:t>D</w:t></w:r></w:p>"#
    ));

    let mut runs = base.edit();
    assert_eq!(
        runs.move_run(Position::new(0), Position::new(0), RunAnchor::End)
            .unwrap(),
        Position::new(1)
    );
    assert_eq!(
        runs.projected()
            .paragraph(Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "BA"
    );
    assert!(
        std::str::from_utf8(runs.projected().xml_bytes())
            .unwrap()
            .contains(r#"<w:opaque w:value="keep"/>"#)
    );
    let moved = commit_and_check(&base, runs);
    assert!(
        std::str::from_utf8(moved.snapshot().xml_bytes())
            .unwrap()
            .contains("<!-- run-owned -->")
    );

    let mut cross = base.edit();
    assert_eq!(
        cross
            .move_run_to(
                Position::new(0),
                Position::new(0),
                Position::new(1),
                RunAnchor::End,
            )
            .unwrap(),
        Position::new(1)
    );
    assert_eq!(
        cross
            .projected()
            .paragraph(Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "B"
    );
    assert_eq!(
        cross
            .projected()
            .paragraph(Position::new(1))
            .unwrap()
            .text()
            .unwrap(),
        "CA"
    );
    let cross_commit = commit_and_check(&base, cross);
    let restored = cross_commit
        .patch()
        .inverse()
        .apply(cross_commit.snapshot())
        .unwrap();
    assert_eq!(restored.xml_bytes(), base.xml_bytes());

    let mut swapped = base.edit();
    swapped
        .swap_runs(Position::new(0), Position::new(0), Position::new(1))
        .unwrap();
    swapped
        .swap_paragraphs(Position::new(0), Position::new(2))
        .unwrap();
    assert_eq!(
        texts(swapped.projected()),
        vec!["D".to_owned(), "C".to_owned(), "BA".to_owned()]
    );
    let swapped_commit = commit_and_check(&base, swapped);
    assert!(
        swapped_commit
            .patch()
            .operations()
            .iter()
            .filter(|operation| {
                matches!(
                    operation,
                    Operation::InsertRun { .. } | Operation::RemoveRun { .. }
                )
            })
            .count()
            >= 2
    );
}

#[test]
fn cross_paragraph_run_move_refuses_unproven_namespace_scope() {
    let base = Snapshot::from_xml(format!(
        r#"<w:document xmlns:w="{WORD}" xmlns:x="urn:one"><w:body><w:p><w:r><w:t>A</w:t></w:r></w:p><w:p xmlns:x="urn:two"><w:r><w:t>B</w:t></w:r></w:p></w:body></w:document>"#
    ))
    .unwrap();
    let mut edit = base.edit();
    assert!(matches!(
        edit.move_run_to(
            Position::new(0),
            Position::new(0),
            Position::new(1),
            RunAnchor::End,
        ),
        Err(Error::Refused {
            reason: Refusal::ComplexContent,
            ..
        })
    ));
    assert_eq!(edit.projected().xml_bytes(), base.xml_bytes());
}

#[test]
fn cross_paragraph_run_move_opens_empty_target_and_replays_exactly() {
    let base = snapshot(r#"<w:p><w:r><w:t>moved</w:t><w:opaque w:value="run"/></w:r></w:p><w:p/>"#);
    let mut edit = base.edit();
    assert_eq!(
        edit.move_run_to(
            Position::new(0),
            Position::new(0),
            Position::new(1),
            RunAnchor::End,
        )
        .unwrap(),
        Position::new(0)
    );
    assert_eq!(
        texts(edit.projected()),
        vec!["".to_owned(), "moved".to_owned()]
    );
    assert!(
        std::str::from_utf8(edit.projected().xml_bytes())
            .unwrap()
            .contains(r#"<w:opaque w:value="run"/>"#)
    );
    commit_and_check(&base, edit);
}

#[test]
fn swaps_exchange_physical_slots_around_opaque_children() {
    let table =
        r#"<w:tbl><w:tr><w:tc><w:p><w:r><w:t>table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#;
    let base_xml = format!(
        r#"<w:p><w:r><w:t>A</w:t></w:r></w:p>{table}<w:p><w:r><w:t>B</w:t></w:r></w:p><w:opaque w:id="between"/><w:p><w:r><w:t>C</w:t></w:r></w:p>"#
    );
    let base = snapshot(&base_xml);
    let mut paragraphs = base.edit();
    paragraphs
        .swap_paragraphs(Position::new(0), Position::new(2))
        .unwrap();
    let children = container_children(paragraphs.projected().xml_bytes(), "body");
    assert_eq!(local_name(&children[0]), "p");
    assert_eq!(local_name(&children[1]), "tbl");
    assert_eq!(local_name(&children[2]), "p");
    assert_eq!(local_name(&children[3]), "opaque");
    assert_eq!(local_name(&children[4]), "p");
    assert!(
        children[0]
            .windows(b">C<".len())
            .any(|window| window == b">C<")
    );
    assert!(
        children[2]
            .windows(b">B<".len())
            .any(|window| window == b">B<")
    );
    assert!(
        children[4]
            .windows(b">A<".len())
            .any(|window| window == b">A<")
    );
    assert_eq!(children[1], table.as_bytes());
    assert_eq!(texts(paragraphs.projected()), vec!["C", "B", "A"]);
    commit_and_check(&base, paragraphs);

    let run_base = snapshot(
        r#"<w:p><w:r><w:t>A</w:t></w:r><w:opaque w:id="one"/><w:r><w:t>B</w:t></w:r><w:opaque w:id="two"/><w:r><w:t>C</w:t></w:r></w:p>"#,
    );
    let mut runs = run_base.edit();
    runs.swap_runs(Position::new(0), Position::new(0), Position::new(2))
        .unwrap();
    let paragraph = paragraph_xml(runs.projected(), 0);
    let children = container_children(&paragraph, "p");
    assert_eq!(local_name(&children[0]), "r");
    assert_eq!(local_name(&children[1]), "opaque");
    assert_eq!(local_name(&children[2]), "r");
    assert_eq!(local_name(&children[3]), "opaque");
    assert_eq!(local_name(&children[4]), "r");
    assert!(
        children[0]
            .windows(b">C<".len())
            .any(|window| window == b">C<")
    );
    assert!(
        children[2]
            .windows(b">B<".len())
            .any(|window| window == b">B<")
    );
    assert!(
        children[4]
            .windows(b">A<".len())
            .any(|window| window == b">A<")
    );
    commit_and_check(&run_base, runs);
}

#[test]
fn cross_paragraph_run_swap_keeps_each_opaque_slot() {
    let base = snapshot(
        r#"<w:p><w:r><w:t>A</w:t></w:r><w:opaque w:id="left"/><w:r><w:t>B</w:t></w:r></w:p><w:p><w:opaque w:id="before"/><w:r><w:t>C</w:t></w:r><w:opaque w:id="after"/><w:r><w:t>D</w:t></w:r></w:p>"#,
    );
    let mut edit = base.edit();
    edit.swap_runs_to(
        Position::new(0),
        Position::new(0),
        Position::new(1),
        Position::new(0),
    )
    .unwrap();
    assert_eq!(texts(edit.projected()), vec!["CB", "AD"]);
    let first = paragraph_xml(edit.projected(), 0);
    let second = paragraph_xml(edit.projected(), 1);
    let first_children = container_children(&first, "p");
    let second_children = container_children(&second, "p");
    assert_eq!(local_name(&first_children[0]), "r");
    assert_eq!(local_name(&first_children[1]), "opaque");
    assert_eq!(local_name(&first_children[2]), "r");
    assert_eq!(local_name(&second_children[0]), "opaque");
    assert_eq!(local_name(&second_children[1]), "r");
    assert_eq!(local_name(&second_children[2]), "opaque");
    assert_eq!(local_name(&second_children[3]), "r");
    assert!(
        first_children[0]
            .windows(b">C<".len())
            .any(|window| window == b">C<")
    );
    assert!(
        second_children[1]
            .windows(b">A<".len())
            .any(|window| window == b">A<")
    );
    commit_and_check(&base, edit);
}

#[test]
fn run_structure_refuses_field_interiors_and_story_anchors() {
    let base = snapshot(concat!(
        r#"<w:p><w:r><w:t>a</w:t></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        r#"<w:r><w:footnoteReference w:id="2"/></w:r></w:p>"#,
        r#"<w:p><w:r><w:t>result tail</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#
    ));
    let mut edit = base.edit();
    for run in [2, 3, 4, 5] {
        assert_eq!(
            refusal(edit.insert_run(Position::new(0), Position::new(run), &NewRun::new("x"))),
            Refusal::FieldInterior,
            "run {run}"
        );
    }
    // A field that began in an earlier paragraph and ends here.
    assert_eq!(
        refusal(edit.insert_run(Position::new(1), Position::new(0), &NewRun::new("x"))),
        Refusal::FieldInterior
    );
    for (run, expected) in [
        (1, Refusal::OpenRange),
        (2, Refusal::OpenRange),
        (6, Refusal::StoryReference),
    ] {
        assert_eq!(
            refusal(edit.remove_run(Position::new(0), Position::new(run))),
            expected
        );
    }
    edit.insert_run(
        Position::new(0),
        Position::new(1),
        &NewRun::new("before field"),
    )
    .unwrap();
    edit.insert_run(
        Position::new(0),
        Position::new(7),
        &NewRun::new("after field"),
    )
    .unwrap();
    edit.append_run(Position::new(1), &NewRun::new("after the end"))
        .unwrap();
    commit_and_check(&base, edit);
}

#[test]
fn new_runs_refuse_text_xml_cannot_carry() {
    let base = snapshot(r#"<w:p/>"#);
    let mut edit = base.edit();
    assert_eq!(
        refusal(edit.append_run(Position::new(0), &NewRun::new("bad\u{0}"))),
        Refusal::StructuralText
    );
    let mut paragraph = NewParagraph::new();
    paragraph.properties_mut().outline_level = Some(12);
    assert_eq!(
        refusal(edit.insert_formatted_paragraph(BodyAnchor::End, &paragraph)),
        Refusal::InvalidPropertyValue
    );
    assert!(edit.operations.is_empty());
}

// ---------------------------------------------------------------------------
// Composition, layout and the corpus.
// ---------------------------------------------------------------------------

#[test]
fn formatting_composes_with_disjoint_edits_and_conflicts_with_overlapping_ones() {
    let base = snapshot(concat!(
        r#"<w:p><w:r><w:t>one</w:t></w:r><w:r><w:t>two</w:t></w:r></w:p>"#,
        r#"<w:p><w:ins w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:t>new</w:t></w:r></w:ins></w:p>"#
    ));
    let limits = CompositionLimits::new(16, 4_096, 4_096, 64);
    let prepared = |id: &str, stage: &dyn Fn(&mut Edit)| {
        let mut edit = base.edit();
        stage(&mut edit);
        edit.prepare(limits, id).unwrap()
    };
    let bold = prepared("a", &|edit| {
        edit.format_run(Position::new(0), Position::new(0), |run| {
            run.bold = Some(true)
        })
        .unwrap();
    });
    let italic = prepared("b", &|edit| {
        edit.format_run(Position::new(0), Position::new(1), |run| {
            run.italic = Some(true)
        })
        .unwrap();
    });
    let text = prepared("c", &|edit| {
        edit.replace_run_text(Position::new(0), Position::new(0), "ONE")
            .unwrap();
    });
    let paragraph = prepared("d", &|edit| {
        edit.format_paragraph(Position::new(0), |paragraph| {
            paragraph.keep_next = Some(true);
        })
        .unwrap();
    });
    let mut composition = base.compose(limits);
    composition.join(bold).unwrap();
    composition.join(italic).unwrap();
    composition.join(text).unwrap();
    composition.join(paragraph).unwrap();
    let composed = commit_composition_for_test(composition);
    let mut sequential = base.edit();
    sequential
        .format_run(Position::new(0), Position::new(0), |run| {
            run.bold = Some(true)
        })
        .unwrap();
    sequential
        .format_run(Position::new(0), Position::new(1), |run| {
            run.italic = Some(true)
        })
        .unwrap();
    sequential
        .replace_run_text(Position::new(0), Position::new(0), "ONE")
        .unwrap();
    sequential
        .format_paragraph(Position::new(0), |paragraph| {
            paragraph.keep_next = Some(true)
        })
        .unwrap();
    assert_eq!(
        composed.snapshot().xml_bytes(),
        sequential.commit().unwrap().snapshot().xml_bytes()
    );

    // Same run property, a run insertion, a structural paragraph edit and a
    // revision disposition all overlap a run property change.
    let conflicts: [&dyn Fn(&mut Edit); 4] = [
        &|edit| {
            edit.format_run(Position::new(0), Position::new(0), |run| {
                run.caps = Some(true)
            })
            .unwrap();
        },
        &|edit| {
            edit.append_run(Position::new(0), &NewRun::new("x"))
                .unwrap();
        },
        &|edit| {
            edit.insert_formatted_paragraph(BodyAnchor::End, &NewParagraph::new())
                .unwrap();
        },
        &|edit| {
            edit.replace_paragraph_text(Position::new(0), "whole")
                .unwrap();
        },
    ];
    for (index, conflict) in conflicts.iter().enumerate() {
        let mut composition = base.compose(limits);
        composition
            .join(prepared("a", &|edit| {
                edit.format_run(Position::new(0), Position::new(0), |run| {
                    run.bold = Some(true);
                })
                .unwrap();
            }))
            .unwrap();
        assert!(
            composition.join(prepared("b", *conflict)).is_err(),
            "conflict {index} joined"
        );
    }
    let mut composition = base.compose(limits);
    composition
        .join(prepared("a", &|edit| {
            edit.format_paragraph(Position::new(1), |paragraph| {
                paragraph.keep_lines = Some(true);
            })
            .unwrap();
        }))
        .unwrap();
    assert!(
        composition
            .join(prepared("b", &|edit| {
                edit.accept_revision(Position::new(1), RevisionKind::Insertion, Position::new(0))
                    .unwrap();
            }))
            .is_err()
    );
}

#[test]
fn run_and_property_rewrites_keep_a_derived_layout() {
    let base = snapshot(concat!(
        r#"<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>"#,
        r#"<w:p><w:r><w:t>a</w:t></w:r></w:p>"#,
        r#"<w:sdt><w:sdtContent><w:p/></w:sdtContent></w:sdt>"#,
        r#"<w:p><w:r><w:t>b</w:t></w:r></w:p><w:sectPr/>"#
    ));
    let (start, end) = {
        let range = base.paragraphs[0];
        let start = usize::try_from(range.start).unwrap();
        (start, start + usize::try_from(range.length).unwrap())
    };
    let grown =
        br#"<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>a</w:t></w:r></w:p>"#;
    let xml = replace_range(base.xml_bytes(), start, end, grown).unwrap();
    let splice = [ParagraphSplice {
        start,
        end,
        replacement: grown,
    }];
    assert!(
        base.spliced_layout_under(xml.len(), &splice, SplicePolicy::SameShape)
            .unwrap()
            .is_none()
    );
    assert!(
        base.spliced_layout_under(xml.len(), &splice, SplicePolicy::BoundedGrowth)
            .unwrap()
            .is_some()
    );
    // A grown fragment that names `body`, changes its root tag, or declares
    // a namespace still takes the rescan.
    for replacement in [
        &br#"<w:p><w:body/></w:p>"#[..],
        br#"<w:p w:rsidR="1"><w:r/></w:p>"#,
        br#"<w:p><w:r xmlns:y="urn:y"/></w:p>"#,
    ] {
        let xml = replace_range(base.xml_bytes(), start, end, replacement).unwrap();
        assert!(
            base.spliced_layout_under(
                xml.len(),
                &[ParagraphSplice {
                    start,
                    end,
                    replacement
                }],
                SplicePolicy::BoundedGrowth
            )
            .unwrap()
            .is_none()
        );
    }
    let mut edit = base.edit();
    for paragraph in [0, 1] {
        edit.format_paragraph(Position::new(paragraph), |properties| {
            properties.keep_next = Some(true);
        })
        .unwrap();
        edit.format_run(Position::new(paragraph), Position::new(0), |run| {
            run.size = Some(HalfPoints(30));
        })
        .unwrap();
        edit.append_run(Position::new(paragraph), &NewRun::new("tail"))
            .unwrap();
        edit.remove_run(Position::new(paragraph), Position::new(0))
            .unwrap();
        assert_layout_matches_rescan(edit.projected());
    }
    commit_and_check(&base, edit);
}

#[test]
fn a_thousand_property_sets_in_one_edit_stay_exact() {
    let mut body = String::new();
    for index in 0..50 {
        body.push_str(&format!(
            r#"<w:p><w:r><w:t>run {index}</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>second</w:t></w:r></w:p>"#
        ));
    }
    let base = snapshot(&body);
    let mut edit = base.edit();
    for round in 0..10 {
        for paragraph in 0..50 {
            edit.format_run(Position::new(paragraph), Position::new(round % 2), |run| {
                run.size = Some(HalfPoints(20 + i64::try_from(round).unwrap()));
                run.bold = Some(round % 3 == 0);
            })
            .unwrap();
        }
    }
    assert_eq!(edit.operations.len(), 500);
    let commit = commit_and_check(&base, edit);
    assert_eq!(
        run_properties(commit.snapshot(), 7, 1).size,
        Some(HalfPoints(29))
    );
}

fn docx_fixtures() -> Vec<std::path::PathBuf> {
    let mut paths = Vec::new();
    let mut stack = vec![
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("..")
            .join("..")
            .join("test-data"),
    ];
    while let Some(directory) = stack.pop() {
        let Ok(entries) = std::fs::read_dir(&directory) else {
            continue;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                stack.push(path);
            } else if path
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("docx"))
            {
                paths.push(path);
            }
        }
    }
    paths.sort();
    paths
}

/// Every typed property, set on sampled runs and paragraphs of every DOCX
/// fixture, reads back as set, leaves every other child of its container
/// byte for byte where it was, changes nothing outside the container, keeps
/// a layout equal to a rescan, and has an exact, replayable inverse.
#[test]
fn every_property_on_the_docx_fixture_corpus_reads_back_and_preserves_the_rest() {
    let mut fixtures = 0usize;
    let mut run_edits = 0usize;
    let mut paragraph_edits = 0usize;
    let mut tracked = 0usize;
    for path in docx_fixtures() {
        let Ok(archive) = std::fs::read(&path) else {
            continue;
        };
        let Ok(package) = crate::Package::from_reader(std::io::Cursor::new(archive)) else {
            continue;
        };
        let Ok(base) = package.document_snapshot() else {
            continue;
        };
        if std::str::from_utf8(base.xml_bytes()).is_err() {
            continue;
        }
        fixtures += 1;
        let count = base.paragraph_count();
        let step = count.div_ceil(4).max(1);
        // Besides an even sample, every paragraph (up to eight per fixture)
        // whose properties or runs record tracked changes or carry markup
        // of other namespaces.
        let mut sample = (0..count).step_by(step).collect::<Vec<_>>();
        sample.extend(
            (0..count)
                .filter(|paragraph| {
                    let xml = paragraph_xml(&base, *paragraph);
                    [&b"Change "[..], b"<w:ins ", b"<w:del ", b"<w14:", b"<mc:"]
                        .iter()
                        .any(|marker| find(&xml, marker).is_some())
                })
                .take(8),
        );
        sample.sort_unstable();
        sample.dedup();
        for paragraph in sample {
            let paragraph_before = paragraph_xml(&base, paragraph);
            tracked += usize::from(
                [&b"Change "[..], b"<w:ins ", b"<w:del "]
                    .iter()
                    .any(|marker| find(&paragraph_before, marker).is_some()),
            );
            for (name, changed, set) in paragraph_setters() {
                let mut edit = base.edit();
                if edit
                    .format_paragraph(Position::new(paragraph), set)
                    .is_err()
                {
                    continue;
                }
                let projected = edit.projected().clone();
                let mut expected = paragraph_properties(&base, paragraph);
                set(&mut expected);
                assert!(
                    crate::properties::same_paragraph(
                        &paragraph_properties(&projected, paragraph),
                        &expected
                    ),
                    "{}: paragraph {paragraph} {name}",
                    path.display()
                );
                let paragraph_after = paragraph_xml(&projected, paragraph);
                assert_preserved(&paragraph_before, &paragraph_after, "pPr", changed);
                assert_outside_unchanged(&base, &projected, paragraph, b"pPr");
                if paragraph_edits % 7 == 0 {
                    commit_and_check(&base, edit);
                } else {
                    assert_layout_matches_rescan(&projected);
                }
                paragraph_edits += 1;
            }
            let runs = formatting::direct_run_count_for_tests(&paragraph_before);
            for run in [0, runs.saturating_sub(1)] {
                if run >= runs {
                    continue;
                }
                let run_before = run_xml(&base, paragraph, run);
                for (name, changed, set) in run_setters() {
                    let mut edit = base.edit();
                    if edit
                        .format_run(Position::new(paragraph), Position::new(run), set)
                        .is_err()
                    {
                        continue;
                    }
                    let projected = edit.projected().clone();
                    let mut expected = run_properties(&base, paragraph, run);
                    set(&mut expected);
                    assert!(
                        crate::properties::same_run(
                            &run_properties(&projected, paragraph, run),
                            &expected
                        ),
                        "{}: paragraph {paragraph} run {run} {name}",
                        path.display()
                    );
                    let run_after = run_xml(&projected, paragraph, run);
                    assert_preserved(&run_before, &run_after, "rPr", changed);
                    if run_edits % 11 == 0 {
                        commit_and_check(&base, edit);
                    } else {
                        assert_layout_matches_rescan(&projected);
                    }
                    run_edits += 1;
                }
            }
        }
    }
    assert!(fixtures >= 50, "{fixtures} fixtures");
    assert!(
        paragraph_edits >= 2_000 && run_edits >= 2_000,
        "{paragraph_edits} paragraph and {run_edits} run edits"
    );
    assert!(tracked > 0, "the corpus should reach tracked changes");
    eprintln!(
        "corpus: {fixtures} fixtures, {paragraph_edits} paragraph and {run_edits} run edits, {tracked} paragraphs with tracked changes"
    );
}

/// After the paragraph's head (its start tag through its `container`), the
/// paragraph and the document are byte for byte as they were; an empty
/// paragraph gains only its end tag.
fn assert_outside_unchanged(
    base: &Snapshot,
    projected: &Snapshot,
    paragraph: usize,
    container: &[u8],
) {
    let before = paragraph_xml(base, paragraph);
    let after = paragraph_xml(projected, paragraph);
    let before_rest = &before[formatting::head_len_for_tests(&before, b"p", container)..];
    let after_rest = &after[formatting::head_len_for_tests(&after, b"p", container)..];
    if before.ends_with(b"/>") && before_rest.is_empty() {
        assert_eq!(after_rest, b"</w:p>");
    } else {
        assert_eq!(before_rest, after_rest);
    }
    let range = base.paragraphs[paragraph];
    let start = usize::try_from(range.start).unwrap();
    let end = start + usize::try_from(range.length).unwrap();
    assert_eq!(&base.xml_bytes()[..start], &projected.xml_bytes()[..start]);
    assert_eq!(
        &base.xml_bytes()[end..],
        &projected.xml_bytes()[start + after.len()..]
    );
}

/// A commit of these operations checks the paragraphs they write, not the
/// whole part, so a paragraph that is not UTF-8 is refused when it is
/// staged, and an untouched one elsewhere is left to the publication audit.
#[test]
fn a_paragraph_that_is_not_utf8_is_refused_when_staged() {
    let mut xml =
        document(r#"<w:p><w:r><w:t>caf_</w:t></w:r></w:p><w:p><w:r><w:t>ok</w:t></w:r></w:p>"#);
    let latin1 = xml.iter().position(|byte| *byte == b'_').unwrap();
    xml[latin1] = 0xE9;
    let base = Snapshot::from_xml(xml).unwrap();
    let mut edit = base.edit();
    assert!(matches!(
        edit.format_run(Position::new(0), Position::new(0), |run| run.bold =
            Some(true)),
        Err(Error::Document(crate::Error::InvalidFormat(_)))
    ));
    assert!(edit.operations.is_empty());
    edit.format_run(Position::new(1), Position::new(0), |run| {
        run.bold = Some(true)
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.patch().is_noop());
}

#[test]
fn paragraph_insertion_and_removal_keep_a_derived_layout() {
    let base = snapshot(BLOCKS);
    let body = formatting::body_for_tests(&base.edit());
    let fragment = b"<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:t>new</w:t></w:r></w:p>";
    // Before the table: offset of the table, paragraph position 1.
    let offset = body[1];
    let xml = replace_range(base.xml_bytes(), offset, offset, fragment).unwrap();
    assert!(
        base.inserted_layout(xml.len(), offset, 1, fragment, fragment.len())
            .unwrap()
            .is_some()
    );
    let derived = base
        .with_inserted_paragraph(xml, offset, 1, fragment, fragment.len())
        .unwrap();
    assert_layout_matches_rescan(&derived);
    assert_eq!(derived.paragraph_count(), 4);
    // A fragment that declares a namespace, holds a processing instruction,
    // is not a paragraph, or uses an unbound prefix takes the rescan.
    for fragment in [
        &br#"<w:p xmlns:y="urn:y"/>"#[..],
        b"<w:p><?pi?></w:p>",
        b"<w:tbl/>",
        b"<q:p/>",
        b"<w:p/><w:p/>",
    ] {
        let xml = replace_range(base.xml_bytes(), offset, offset, fragment).unwrap();
        assert!(
            base.inserted_layout(xml.len(), offset, 1, fragment, fragment.len())
                .unwrap()
                .is_none(),
            "{}",
            String::from_utf8_lossy(fragment)
        );
    }
    // An offset inside a table, or a position that disagrees with it.
    let inside = offset + 3;
    let xml = replace_range(base.xml_bytes(), inside, inside, fragment).unwrap();
    assert!(
        base.inserted_layout(xml.len(), inside, 1, fragment, fragment.len())
            .unwrap()
            .is_none()
    );
    let xml = replace_range(base.xml_bytes(), offset, offset, fragment).unwrap();
    assert!(
        base.inserted_layout(xml.len(), offset, 0, fragment, fragment.len())
            .unwrap()
            .is_none()
    );
    // A body without a neighbouring paragraph has no proven scope.
    let empty = snapshot("<w:tbl/>");
    let at = usize::try_from(empty.content_end).unwrap();
    let xml = replace_range(empty.xml_bytes(), at, at, b"<w:p/>").unwrap();
    assert!(
        empty
            .inserted_layout(xml.len(), at, 0, b"<w:p/>", 6)
            .unwrap()
            .is_none()
    );
    // Removal always derives.
    for position in 0..base.paragraph_count() {
        let range = base.paragraphs[position];
        let start = usize::try_from(range.start).unwrap();
        let end = start + usize::try_from(range.length).unwrap();
        let xml = replace_range(base.xml_bytes(), start, end, b"").unwrap();
        assert!(
            base.removed_layout(xml.len(), position, end - start)
                .unwrap()
                .is_some()
        );
        assert_layout_matches_rescan(
            &base
                .with_removed_paragraph(xml, position, end - start)
                .unwrap(),
        );
    }
}

/// Structure operations on every DOCX fixture keep a layout equal to a
/// rescan and have exact, replayable inverses.
#[test]
fn structure_operations_on_the_docx_fixture_corpus_round_trip() {
    let mut fixtures = 0usize;
    let mut operations = 0usize;
    for path in docx_fixtures() {
        let Ok(archive) = std::fs::read(&path) else {
            continue;
        };
        let Ok(package) = crate::Package::from_reader(std::io::Cursor::new(archive)) else {
            continue;
        };
        let Ok(base) = package.document_snapshot() else {
            continue;
        };
        if std::str::from_utf8(base.xml_bytes()).is_err() || base.paragraph_count() == 0 {
            continue;
        }
        fixtures += 1;
        let last = base.paragraph_count() - 1;
        let mut paragraph = NewParagraph::new().with_run(NewRun::new("corpus insertion"));
        paragraph.properties_mut().keep_next = Some(true);
        let mut edit = base.edit();
        for anchor in [
            BodyAnchor::Before(BodyBlock::Paragraph(Position::new(0))),
            BodyAnchor::After(BodyBlock::Paragraph(Position::new(last))),
            BodyAnchor::End,
        ] {
            edit.insert_formatted_paragraph(anchor, &paragraph)
                .unwrap_or_else(|error| panic!("{}: {error}", path.display()));
            assert_layout_matches_rescan(edit.projected());
            operations += 1;
        }
        for position in [1usize, last] {
            if edit
                .move_paragraph(Position::new(position), BodyAnchor::End)
                .is_ok()
            {
                assert_layout_matches_rescan(edit.projected());
                operations += 1;
            }
        }
        if edit.remove_paragraph(Position::new(1)).is_ok() {
            assert_layout_matches_rescan(edit.projected());
            operations += 1;
        }
        if edit
            .append_run(Position::new(0), &NewRun::new("tail"))
            .is_ok()
        {
            operations += 1;
        }
        commit_and_check(&base, edit);
    }
    assert!(fixtures >= 50, "{fixtures} fixtures");
    assert!(operations >= 300, "{operations} operations");
}

/// In an indented document every removed or moved element takes the
/// whitespace that follows it, so durable inverse replay restores the
/// indentation byte for byte.
#[test]
fn indented_documents_round_trip_structure_exactly() {
    let base = snapshot(concat!(
        "\r\n    <w:p>\r\n      <w:r>\r\n        <w:t>one</w:t>\r\n      </w:r>\r\n",
        "      <w:r>\r\n        <w:t>two</w:t>\r\n      </w:r>\r\n    </w:p>",
        "\r\n    <w:bookmarkStart w:id=\"1\" w:name=\"b\"/>",
        "\r\n    <w:p>\r\n      <w:r>\r\n        <w:t>three</w:t>\r\n      </w:r>\r\n    </w:p>",
        "\r\n    <!-- kept with three -->",
        "\r\n    <w:bookmarkEnd w:id=\"1\"/>",
        "\r\n    <w:p/>\r\n    <w:sectPr/>\r\n  "
    ));
    let stages: [fn(&mut Edit); 7] = [
        |edit| {
            edit.remove_run(Position::new(0), Position::new(0)).unwrap();
        },
        |edit| {
            edit.remove_run(Position::new(0), Position::new(1)).unwrap();
        },
        |edit| {
            edit.remove_paragraph(Position::new(1)).unwrap();
        },
        |edit| {
            edit.remove_paragraph(Position::new(2)).unwrap();
        },
        |edit| {
            edit.move_paragraph(Position::new(1), BodyAnchor::End)
                .unwrap();
        },
        |edit| {
            edit.move_paragraph(
                Position::new(2),
                BodyAnchor::Before(BodyBlock::Paragraph(Position::new(0))),
            )
            .unwrap();
        },
        |edit| {
            edit.insert_formatted_paragraph(BodyAnchor::End, &NewParagraph::new())
                .unwrap();
            edit.append_run(Position::new(1), &NewRun::new("four"))
                .unwrap();
        },
    ];
    for stage in stages {
        let mut edit = base.edit();
        stage(&mut edit);
        commit_and_check(&base, edit);
    }
    // The comment after paragraph "three" moves with it.
    let mut edit = base.edit();
    edit.move_paragraph(Position::new(1), BodyAnchor::End)
        .unwrap();
    let moved = commit_and_check(&base, edit);
    let xml = std::str::from_utf8(moved.snapshot().xml_bytes()).unwrap();
    assert!(xml.contains("three</w:t>\r\n      </w:r>\r\n    </w:p>\r\n    <!-- kept with three -->\r\n    <w:sectPr/>"));
}

#[test]
fn formatting_resolves_inherited_aliases_and_keeps_foreign_properties() {
    let base = Snapshot::from_xml(format!(r#"<w:document xmlns:w="{WORD}" xmlns:q="{WORD}" xmlns:f="urn:foreign"><w:body><w:p><q:pPr><q:jc q:val="center"/></q:pPr><q:r><q:rPr><q:i/><f:b f:val="PRIVATE"/></q:rPr><q:t>x</q:t></q:r><f:r/></w:p></w:body></w:document>"#)).unwrap();
    let mut edit = base.edit();
    edit.format_run(Position::new(0), Position::new(0), |run| {
        assert_eq!(run.italic, Some(true));
        run.bold = Some(true);
    })
    .unwrap();
    edit.format_paragraph(Position::new(0), |paragraph| {
        assert_eq!(paragraph.justification, Some(Justification::Center));
        paragraph.justification = Some(Justification::Right);
    })
    .unwrap();
    let commit = commit_and_check(&base, edit);
    let xml = std::str::from_utf8(commit.snapshot().xml_bytes()).unwrap();
    assert!(xml.contains(r#"<f:b f:val="PRIVATE"/>"#));
    assert!(xml.contains("<q:b/>"));
    assert!(xml.contains(r#"q:val="right""#));
}

#[test]
fn property_container_removal_preserves_opaque_non_element_content() {
    for content in ["<!--PRIVATE-->", "PRIVATE", "<![CDATA[PRIVATE]]>"] {
        let base = snapshot(&format!(
            "<w:p><w:pPr><w:tabs>{content}</w:tabs></w:pPr></w:p>"
        ));
        let mut edit = base.edit();
        assert!(matches!(
            edit.format_paragraph(Position::new(0), |paragraph| paragraph.tabs = None),
            Err(Error::Refused {
                reason: Refusal::ComplexContent,
                ..
            })
        ));
        assert_eq!(edit.projected().xml_bytes(), base.xml_bytes());
        assert!(edit.operations.is_empty());
    }
}

#[test]
fn foreign_range_ids_cannot_authorize_paragraph_removal() {
    let base = snapshot(r#"<w:p><w:bookmarkStart x:id="1"/><w:bookmarkEnd x:id="1"/></w:p>"#);
    let mut edit = base.edit();
    assert!(edit.remove_paragraph(Position::new(0)).is_err());
    assert_eq!(edit.projected().xml_bytes(), base.xml_bytes());
}

#[test]
fn replay_refuses_duplicate_attributes_before_publishing_fragments() {
    let base = snapshot("<w:p><w:r><w:t>x</w:t></w:r></w:p>");
    let mut edit = base.edit();
    let run = Arc::new(br#"<w:r><w:rPr><w:b w:val="1" w:val="0"/></w:rPr></w:r>"#.to_vec());
    assert!(
        edit.insert_run_fragment(Position::new(0), Position::new(0), Position::new(0), &run)
            .is_err()
    );
    let paragraph = Arc::new(br#"<w:p><w:r PRIVATE="1" PRIVATE="2"/></w:p>"#.to_vec());
    assert!(
        edit.insert_paragraph_fragment(Position::new(0), Position::new(0), &paragraph)
            .is_err()
    );
    assert_eq!(edit.projected().xml_bytes(), base.xml_bytes());
    assert!(edit.operations.is_empty());
}

#[test]
fn formatting_staging_limits_are_typed_and_atomic() {
    let base = snapshot("<w:p><w:r><w:t>x</w:t></w:r></w:p>");
    let mut edit = base.edit();
    let result = edit.format_run(Position::new(0), Position::new(0), |run| {
        run.style = Some("&".repeat(12 * 1024 * 1024))
    });
    assert!(matches!(
        result,
        Err(Error::Limit {
            kind: crate::document::LimitKind::FormattingStagingBytes,
            ..
        })
    ));
    assert!(edit.operations.is_empty());
    assert_eq!(edit.projected().xml_bytes(), base.xml_bytes());
    assert!(!format!("{:?}", NewRun::new("PRIVATE")).contains("PRIVATE"));
    assert!(
        !format!("{:?}", NewParagraph::new().with_run(NewRun::new("PRIVATE"))).contains("PRIVATE")
    );
}

#[test]
fn unchanged_large_properties_bypass_generation_admission() {
    let xml = format!(
        "<w:p><w:r><w:rPr><w:rStyle w:val=\"{}\"/></w:rPr></w:r></w:p>",
        "x".repeat(11 * 1024 * 1024)
    );
    let base = snapshot(&xml);
    let mut edit = base.edit();
    edit.format_run(Position::new(0), Position::new(0), |_| {})
        .unwrap();
    assert!(edit.operations.is_empty());
    assert_eq!(edit.projected().xml_bytes(), base.xml_bytes());
}
