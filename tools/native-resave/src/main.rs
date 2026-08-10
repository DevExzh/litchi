use std::env;
use std::error::Error;
use std::fs;
use std::path::{Path, PathBuf};

#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "xlsx",
    feature = "pptx",
    feature = "rtf",
    feature = "odt",
    feature = "odp"
))]
const SENTINEL: &str = "Litchi native resave 2026-08-10";
#[cfg(feature = "xls")]
const XLS_SENTINEL: f64 = 42.25;
#[cfg(feature = "odb")]
const ODB_QUERY: &str = "__litchi_native_resave";

type Result<T> = std::result::Result<T, Box<dyn Error>>;

fn main() -> Result<()> {
    let mut arguments = env::args_os().skip(1);
    let action = arguments
        .next()
        .ok_or("expected generate or readback")?
        .into_string()
        .map_err(|_| "action is not UTF-8")?;
    let format = arguments
        .next()
        .ok_or("expected format")?
        .into_string()
        .map_err(|_| "format is not UTF-8")?;
    let path = PathBuf::from(arguments.next().ok_or("expected artifact path")?);
    if arguments.next().is_some() {
        return Err("unexpected trailing argument".into());
    }
    match action.as_str() {
        "generate" => generate(&format, &path),
        "readback" => readback(&format, &path),
        _ => Err(format!("unknown action: {action}").into()),
    }
}

fn generate(format: &str, output: &Path) -> Result<()> {
    match format {
        #[cfg(feature = "doc")]
        "doc" => generate_doc(output),
        #[cfg(feature = "docx")]
        "docx" => generate_docx(output),
        #[cfg(feature = "xlsx")]
        "xlsx" => generate_xlsx(output),
        #[cfg(feature = "pptx")]
        "pptx" => generate_pptx(output),
        #[cfg(feature = "ppt")]
        "ppt" => generate_ppt(output),
        #[cfg(feature = "rtf")]
        "rtf" => generate_rtf(output),
        #[cfg(feature = "xls")]
        "xls" => generate_xls(output),
        #[cfg(feature = "odt")]
        "odt" => generate_odt(output),
        #[cfg(feature = "ods")]
        "ods" => generate_ods(output),
        #[cfg(feature = "odp")]
        "odp" => generate_odp(output),
        #[cfg(feature = "odf")]
        "odf" => generate_odf(output),
        #[cfg(feature = "odb")]
        "odb" => generate_odb(output),
        #[cfg(feature = "odg")]
        "odg" => generate_odg(output),
        _ => Err(format!("unsupported generator format: {format}").into()),
    }
}

fn readback(format: &str, input: &Path) -> Result<()> {
    match format {
        #[cfg(feature = "doc")]
        "doc" => readback_doc(input),
        #[cfg(feature = "docx")]
        "docx" => readback_docx(input),
        #[cfg(feature = "xlsx")]
        "xlsx" => readback_xlsx(input),
        #[cfg(feature = "pptx")]
        "pptx" => readback_pptx(input),
        #[cfg(feature = "ppt")]
        "ppt" => readback_ppt(input),
        #[cfg(feature = "rtf")]
        "rtf" => readback_rtf(input),
        #[cfg(feature = "xls")]
        "xls" => readback_xls(input),
        #[cfg(feature = "odt")]
        "odt" => readback_odt(input),
        #[cfg(feature = "ods")]
        "ods" => readback_ods(input),
        #[cfg(feature = "odp")]
        "odp" => readback_odp(input),
        #[cfg(feature = "odf")]
        "odf" => readback_odf(input),
        #[cfg(feature = "odb")]
        "odb" => readback_odb(input),
        #[cfg(feature = "odg")]
        "odg" => readback_odg(input),
        _ => Err(format!("unsupported readback format: {format}").into()),
    }
}

fn fixture(path: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(path)
}

#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "xlsx",
    feature = "pptx",
    feature = "rtf",
    feature = "xls",
    feature = "odt",
    feature = "ods",
    feature = "odp",
    feature = "odf",
    feature = "odb",
    feature = "odg"
))]
fn missing(message: &str) -> std::io::Error {
    std::io::Error::new(std::io::ErrorKind::NotFound, message)
}

#[cfg(feature = "doc")]
fn generate_doc(output: &Path) -> Result<()> {
    let source = fixture("test-data/ole/doc/NoHeadFoot.doc");
    let snapshot = litchi_doc::body_text::Snapshot::from_bytes(fs::read(source)?)?;
    let paragraph = snapshot
        .paragraphs(litchi_doc::body_text::Projection::All)?
        .into_iter()
        .find(|paragraph| !paragraph.text().is_empty())
        .ok_or_else(|| missing("DOC has no nonempty ordinary paragraph"))?;
    println!("doc source paragraph={}", paragraph.text());
    let mut edit = snapshot.edit()?;
    edit.replace_paragraph(paragraph.position(), SENTINEL)?;
    let commit = edit.commit()?;
    fs::write(output, commit.snapshot().finish())?;
    readback_doc(output)
}

#[cfg(feature = "doc")]
fn readback_doc(input: &Path) -> Result<()> {
    let snapshot = litchi_doc::body_text::Snapshot::from_bytes(fs::read(input)?)?;
    let found = snapshot
        .paragraphs(litchi_doc::body_text::Projection::All)?
        .into_iter()
        .any(|paragraph| paragraph.text() == SENTINEL);
    if !found {
        return Err("DOC sentinel paragraph is absent after reopen".into());
    }
    println!("doc paragraph={SENTINEL}");
    Ok(())
}

#[cfg(feature = "docx")]
fn generate_docx(output: &Path) -> Result<()> {
    let source = fixture("test-data/ooxml/docx/documentProperties.docx");
    let mut package = litchi_docx::Package::open(source)?;
    let position = litchi_core::Position::new(0);
    let before = package
        .document_snapshot()?
        .paragraph(position)
        .ok_or_else(|| missing("DOCX paragraph 0 is absent"))?
        .text()?;
    println!("docx source paragraph[0]={before}");
    let mut edit = package.edit_document()?;
    edit.replace_paragraph_text(position, SENTINEL)?;
    package.publish_document_edit(edit)?;
    package.save(output)?;
    readback_docx(output)
}

#[cfg(feature = "docx")]
fn readback_docx(input: &Path) -> Result<()> {
    let package = litchi_docx::Package::open(input)?;
    let text = package
        .document_snapshot()?
        .paragraph(litchi_core::Position::new(0))
        .ok_or_else(|| missing("DOCX paragraph 0 is absent after reopen"))?
        .text()?;
    if text != SENTINEL {
        return Err(format!("DOCX readback mismatch: {text:?}").into());
    }
    println!("docx paragraph[0]={text}");
    Ok(())
}

#[cfg(feature = "xlsx")]
fn generate_xlsx(output: &Path) -> Result<()> {
    let source = fixture("test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx");
    let workbook = litchi_xlsx::Workbook::open(source)?;
    let sheet = workbook
        .sheets()
        .next()
        .ok_or_else(|| missing("XLSX has no worksheet"))?;
    let name = sheet.name().to_string();
    let mut edit = workbook.edit()?;
    edit.sheet(name.as_str())?
        .ok_or_else(|| missing("XLSX first sheet is not editable"))?
        .set("A1", SENTINEL)?;
    let changed = edit.commit()?.into_workbook();
    changed.save_plain(output)?;
    readback_xlsx(output)
}

#[cfg(feature = "xlsx")]
fn readback_xlsx(input: &Path) -> Result<()> {
    let workbook = litchi_xlsx::Workbook::open(input)?;
    let sheet = workbook
        .sheets()
        .next()
        .ok_or_else(|| missing("XLSX has no worksheet after reopen"))?;
    let stored = sheet
        .cell("A1")?
        .stored()
        .ok_or_else(|| missing("XLSX A1 is absent after reopen"))?;
    match stored {
        litchi_xlsx::Cell::Value(litchi_xlsx::Value::Text(text)) if text.as_str() == SENTINEL => {
            println!("xlsx {}!A1={text}", sheet.name());
            Ok(())
        },
        other => Err(format!("XLSX readback mismatch: {other:?}").into()),
    }
}

#[cfg(feature = "pptx")]
fn generate_pptx(output: &Path) -> Result<()> {
    let source = fixture("test-data/ooxml/pptx/shapes.pptx");
    let bytes = fs::read(source)?;
    let mut package = litchi_pptx::Package::from_bytes(&bytes)?;
    let root = package.opened_presentation()?;
    let slide = package
        .presentation()?
        .slide(0)?
        .ok_or_else(|| missing("PPTX slide 0 is absent"))?;
    let shape = slide
        .shapes()?
        .iter()
        .position(|shape| shape.common().text().is_some())
        .ok_or_else(|| missing("PPTX slide 0 has no text shape"))?;
    let mut edit = root.edit();
    edit.set_shape_text(0, shape, SENTINEL)?;
    let patch = edit.commit()?.into_patch();
    package.apply_opened_presentation_patch(&patch)?;
    fs::write(output, package.to_bytes()?)?;
    readback_pptx(output)
}

#[cfg(feature = "pptx")]
fn readback_pptx(input: &Path) -> Result<()> {
    let bytes = fs::read(input)?;
    let package = litchi_pptx::Package::from_bytes(&bytes)?;
    let slide = package
        .presentation()?
        .slide(0)?
        .ok_or_else(|| missing("PPTX slide 0 is absent after reopen"))?;
    let shapes = slide.shapes()?;
    let text = shapes
        .iter()
        .filter_map(|shape| shape.common().text())
        .find(|text| text.contains(SENTINEL))
        .ok_or_else(|| missing("PPTX sentinel text is absent after reopen"))?;
    println!("pptx slide[0] text={text}");
    Ok(())
}

#[cfg(feature = "ppt")]
fn generate_ppt(output: &Path) -> Result<()> {
    let source = fixture("test-data/poi/test-data/slideshow/45543.ppt");
    let snapshot = litchi_ppt::slide_order::Snapshot::from_bytes(fs::read(source)?)?;
    let position = litchi_core::Position::new(0);
    let before = snapshot.slide_transition_visual(position)?;
    if before != ppt_source_transition()? {
        return Err(format!("PPT source transition mismatch: {before:?}").into());
    }
    let mut edit = snapshot.edit()?;
    edit.set_slide_transition_visual(position, ppt_changed_transition()?)?;
    let commit = edit.commit()?;
    fs::write(output, commit.snapshot().bytes())?;
    readback_ppt(output)
}

#[cfg(feature = "ppt")]
fn readback_ppt(input: &Path) -> Result<()> {
    let position = litchi_core::Position::new(0);
    let changed = litchi_ppt::slide_order::Snapshot::from_bytes(fs::read(input)?)?;
    let visual = changed.slide_transition_visual(position)?;
    if visual != ppt_changed_transition()? {
        return Err(format!("PPT transition mismatch: {visual:?}").into());
    }
    println!(
        "ppt slide[0] transition={:?}; direction={:?}; speed={:?}",
        visual.transition_type(),
        visual.direction(),
        visual.speed()
    );
    Ok(())
}

#[cfg(feature = "ppt")]
fn ppt_source_transition() -> Result<litchi_ppt::slide_order::SlideTransitionVisual> {
    Ok(litchi_ppt::slide_order::SlideTransitionVisual::new(
        litchi_ppt::TransitionType::Box,
        litchi_ppt::TransitionDirection::Out,
        litchi_ppt::TransitionSpeed::Slow,
    )?)
}

#[cfg(feature = "ppt")]
fn ppt_changed_transition() -> Result<litchi_ppt::slide_order::SlideTransitionVisual> {
    Ok(litchi_ppt::slide_order::SlideTransitionVisual::new(
        litchi_ppt::TransitionType::Cover,
        litchi_ppt::TransitionDirection::FromLeft,
        litchi_ppt::TransitionSpeed::Medium,
    )?)
}

#[cfg(feature = "rtf")]
fn generate_rtf(output: &Path) -> Result<()> {
    let source = fs::read(fixture(
        "test-data/libreoffice-core/sw/qa/extras/rtfexport/data/relsize.rtf",
    ))?;
    let document = litchi_rtf::Document::from_bytes(&source)?;
    let mut edit = document.edit();
    edit.set_shape_text(0, SENTINEL)?;
    let changed = edit.commit()?.into_snapshot().to_bytes()?;
    fs::write(output, changed)?;
    readback_rtf(output)
}

#[cfg(feature = "rtf")]
fn readback_rtf(input: &Path) -> Result<()> {
    let bytes = fs::read(input)?;
    let document = litchi_rtf::Document::from_bytes(&bytes)?;
    let text = document
        .shapes()
        .first()
        .ok_or_else(|| missing("RTF shape 0 is absent after reopen"))?
        .text
        .as_ref();
    if text.trim_end() != SENTINEL {
        return Err(format!("RTF readback mismatch: {text:?}").into());
    }
    println!("rtf shape[0]={text}");
    Ok(())
}

#[cfg(feature = "xls")]
fn generate_xls(output: &Path) -> Result<()> {
    let source = fixture("test-data/libreoffice-core/sc/qa/extras/testdocuments/tdf78897.xls");
    let snapshot = litchi_xls::cell_values::Snapshot::from_bytes(fs::read(source)?)?;
    let mut candidates = Vec::new();
    for sheet in snapshot.worksheets() {
        for cell in sheet.cells() {
            if matches!(cell.value(), litchi_xls::cell_values::Value::Number(_)) {
                candidates.push((sheet.name().to_string(), cell.reference()));
            }
        }
    }
    let mut changed = None;
    for (sheet, reference) in candidates {
        let mut edit = snapshot.edit();
        if edit
            .set_value(
                litchi_xls::cell_values::Selector::Name(&sheet),
                reference,
                litchi_xls::cell_values::Value::Number(XLS_SENTINEL),
            )
            .is_err()
        {
            continue;
        }
        let Ok(commit) = edit.commit() else {
            continue;
        };
        println!(
            "xls source target={}!R{}C{}",
            sheet,
            reference.row(),
            reference.column()
        );
        changed = Some(commit.snapshot().bytes().to_vec());
        break;
    }
    fs::write(
        output,
        changed.ok_or_else(|| missing("XLS has no safely editable numeric cell"))?,
    )?;
    readback_xls(output)
}

#[cfg(feature = "xls")]
fn readback_xls(input: &Path) -> Result<()> {
    let snapshot = litchi_xls::cell_values::Snapshot::from_bytes(fs::read(input)?)?;
    let found = snapshot.worksheets().any(|sheet| {
        sheet.cells().any(|cell| {
            matches!(
                cell.value(),
                litchi_xls::cell_values::Value::Number(value) if value.to_bits() == XLS_SENTINEL.to_bits()
            )
        })
    });
    if !found {
        return Err(format!("XLS sentinel number {XLS_SENTINEL} is absent after reopen").into());
    }
    println!("xls numeric sentinel={XLS_SENTINEL}");
    Ok(())
}

#[cfg(feature = "odt")]
fn generate_odt(output: &Path) -> Result<()> {
    let source = fixture("test-data/odf/corpus/writer-header-footer.odt");
    let document = litchi_odt::Document::open(source)?;
    let mut edit = document.edit()?;
    edit.replace_paragraph(litchi_core::Position::new(0), SENTINEL)?;
    let commit = edit.commit()?;
    fs::write(output, commit.snapshot().as_bytes())?;
    readback_odt(output)
}

#[cfg(feature = "odt")]
fn readback_odt(input: &Path) -> Result<()> {
    let document = litchi_odt::Document::open(input)?;
    let text = document.text()?;
    if !text.contains(SENTINEL) {
        return Err("ODT sentinel text is absent after reopen".into());
    }
    println!("odt text contains={SENTINEL}");
    Ok(())
}

#[cfg(feature = "ods")]
fn generate_ods(output: &Path) -> Result<()> {
    let source = fixture("test-data/odf/corpus/calc-two-sheets.ods");
    let bytes = fs::read(source)?;
    let spreadsheet = litchi_ods::Spreadsheet::from_bytes(bytes.clone())?;
    let sheet = spreadsheet
        .sheets()
        .first()
        .ok_or_else(|| missing("ODS has no sheet"))?
        .name
        .to_string();
    let snapshot = litchi_ods::document::Snapshot::from_bytes(bytes)?;
    let mut edit = snapshot.edit();
    edit.set_cell_formula(&sheet, 0, 0, "of:=40+2")?;
    let commit = edit.commit()?;
    fs::write(output, commit.snapshot().as_bytes())?;
    readback_ods(output)
}

#[cfg(feature = "ods")]
fn readback_ods(input: &Path) -> Result<()> {
    let spreadsheet = litchi_ods::Spreadsheet::from_bytes(fs::read(input)?)?;
    if !spreadsheet
        .content_xml()
        .contains("table:formula=\"of:=40+2\"")
    {
        return Err("ODS formula is absent after reopen".into());
    }
    println!("ods formula=of:=40+2");
    Ok(())
}

#[cfg(feature = "odp")]
fn generate_odp(output: &Path) -> Result<()> {
    let source = fixture("test-data/odf/odp/tdf169979.odp");
    let snapshot = litchi_odp::edit::Snapshot::open(source)?;
    let mut edit = snapshot.transaction()?;
    let rich = litchi_odp::content::RichText::plain(SENTINEL)?;
    let text_box = litchi_odp::content::TextBox::new("Litchi Interop Box", rich)?;
    edit.add_text_box(0_usize, &text_box)?;
    let commit = edit.commit()?;
    fs::write(output, commit.snapshot().bytes())?;
    readback_odp(output)
}

#[cfg(feature = "odp")]
fn readback_odp(input: &Path) -> Result<()> {
    let presentation = litchi_odp::Presentation::from_bytes(fs::read(input)?)?;
    let xml = presentation.content_xml();
    if !xml.contains("Litchi Interop Box") || !xml.contains(SENTINEL) {
        return Err("ODP text box is absent after reopen".into());
    }
    println!("odp text-box=Litchi Interop Box; text={SENTINEL}");
    Ok(())
}

#[cfg(feature = "odf")]
fn generate_odf(output: &Path) -> Result<()> {
    let source = fixture("test-data/odf/native-resave/source/font-styles.odf");
    let formula = litchi_odf_formula::Formula::from_bytes(fs::read(source)?)?;
    let mut edit = formula.edit();
    edit.set_text(&litchi_odf_formula::NodePath::new([0, 0, 0]), "g")?;
    edit.set_starmath(&litchi_odf_formula::OpaqueStarMath::new(
        litchi_odf_formula::StarMathVersion::V5,
        "g",
    )?)?;
    let changed = edit.commit()?.into_formula();
    fs::write(output, changed.to_bytes())?;
    readback_odf(output)
}

#[cfg(feature = "odf")]
fn readback_odf(input: &Path) -> Result<()> {
    let formula = litchi_odf_formula::Formula::from_bytes(fs::read(input)?)?;
    let xml = litchi_odf_formula::codec::serialize(formula.root());
    if !xml.contains(">g<") {
        return Err("ODF Formula identifier g is absent after reopen".into());
    }
    let starmath = formula
        .starmath()
        .ok_or_else(|| missing("ODF Formula StarMath annotation is absent"))?;
    if starmath.opaque().source() != "g" {
        return Err("ODF Formula StarMath source g is absent after reopen".into());
    }
    println!("odf token-path[0,0,0]=g; starmath=g");
    Ok(())
}

#[cfg(feature = "odb")]
fn generate_odb(output: &Path) -> Result<()> {
    let source = fixture("test-data/libreoffice-core/dbaccess/qa/unit/data/tdf132924.odb");
    let database = litchi_odb::Database::open(source)?;
    let mut edit = database.edit();
    edit.add_query(litchi_odb::Query::new(ODB_QUERY, "SELECT 424242"))?;
    let commit = edit.commit()?;
    fs::write(output, commit.database().as_bytes())?;
    readback_odb(output)
}

#[cfg(feature = "odb")]
fn readback_odb(input: &Path) -> Result<()> {
    let database = litchi_odb::Database::open(input)?;
    let catalog = database.catalog()?;
    let query = catalog
        .query(ODB_QUERY)?
        .ok_or_else(|| missing("ODB sentinel query is absent after reopen"))?;
    if query.command() != "SELECT 424242" || !query.columns().is_empty() {
        return Err(format!(
            "ODB sentinel query mismatch: command={:?}, columns={}",
            query.command(),
            query.columns().len()
        )
        .into());
    }
    if catalog.table("test")?.is_none() {
        return Err("ODB source table is absent after reopen".into());
    }
    println!("odb query={ODB_QUERY}; command=SELECT 424242; columns=0; source-table=test");
    Ok(())
}

#[cfg(feature = "odg")]
fn generate_odg(output: &Path) -> Result<()> {
    let source = fixture("test-data/odf/native-resave/source/rhbz1870501.odg");
    let drawing = litchi_odg::Drawing::from_bytes(fs::read(source)?)?;
    let group = drawing.pages()[0]
        .shapes()
        .iter()
        .position(|shape| shape.kind() == litchi_odg::shape::ShapeKind::Group)
        .ok_or_else(|| missing("ODG group is absent"))?;
    let descendant = *drawing
        .group(0, group)?
        .descendants()
        .iter()
        .find(|&&index| {
            let shape = &drawing.pages()[0].shapes()[index];
            shape.x().is_some()
                && shape.y().is_some()
                && shape.width().is_some()
                && shape.height().is_some()
        })
        .ok_or_else(|| missing("ODG positioned group descendant is absent"))?;
    let shape = &drawing.pages()[0].shapes()[descendant];
    let y = shape
        .y()
        .ok_or_else(|| missing("ODG descendant y is absent"))?;
    let width = shape
        .width()
        .ok_or_else(|| missing("ODG descendant width is absent"))?;
    let height = shape
        .height()
        .ok_or_else(|| missing("ODG descendant height is absent"))?;
    let mut edit = drawing.edit();
    edit.set_group_descendant_geometry(0, group, descendant, "9cm", y, width, height)?;
    let commit = edit.commit()?;
    fs::write(output, commit.snapshot().as_bytes())?;
    readback_odg(output)
}

#[cfg(feature = "odg")]
fn readback_odg(input: &Path) -> Result<()> {
    let drawing = litchi_odg::Drawing::from_bytes(fs::read(input)?)?;
    if !drawing
        .pages()
        .iter()
        .any(|page| page.shapes().iter().any(|shape| shape.x() == Some("9cm")))
    {
        return Err("ODG x=9cm geometry is absent after reopen".into());
    }
    println!("odg descendant x=9cm");
    Ok(())
}
