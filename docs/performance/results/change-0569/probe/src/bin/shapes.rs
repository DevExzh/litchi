//! Behaviour probe: what does each facade opener do with a mismatched input?
use std::path::Path;

const ROOT: &str = "/home/zhuhe/code/litchi/";

fn show<T, E: std::fmt::Debug + std::fmt::Display>(r: Result<T, E>) -> String {
    match r {
        Ok(_) => "OK".to_owned(),
        Err(e) => {
            let d = format!("{e:?}");
            let s = format!("{e}");
            let d = if d.len() > 150 { format!("{}…", &d[..150]) } else { d };
            let s = if s.len() > 150 { format!("{}…", &s[..150]) } else { s };
            format!("Err[{d}] :: \"{s}\"")
        },
    }
}

fn main() {
    let fixtures: Vec<(&str, &str)> = vec![
        ("DOCX", "test-data/ooxml/docx/Hyperlink.docx"),
        ("PPTX", "test-data/ooxml/pptx/shapes.pptx"),
        ("XLSX", "test-data/ooxml/xlsx/sheet-names.xlsx"),
        ("XLSB", "test-data/ooxml/xlsb/"),
        ("DOC(ole2)", "test-data/ole/doc/NoHeadFoot.doc"),
        ("PPT(ole2)", "test-data/ole/ppt/SampleShow.ppt"),
        ("XLS(ole2)", "test-data/ole/xls/SimpleChart.xls"),
        ("ODT", "test-data/odf/corpus/writer-paragraph-styles.odt"),
        ("ODS", "test-data/odf/corpus/calc-formulas.ods"),
        ("ODP", "test-data/odf/corpus/impress-basic.odp"),
        ("PNG(non-office)", "test-data/images/png/lena.png"),
        ("CSV(non-office)", "test-data/textual/csv/numberformat.csv"),
    ];

    for (label, rel) in fixtures {
        let mut path = format!("{ROOT}{rel}");
        if rel.ends_with('/') {
            // pick the first file in the directory
            let Ok(rd) = std::fs::read_dir(&path) else { continue };
            let mut first: Option<String> = None;
            for e in rd.flatten() {
                let p = e.path();
                if p.is_file() {
                    first = Some(p.to_string_lossy().into_owned());
                    break;
                }
            }
            let Some(f) = first else { continue };
            path = f;
        }
        if !Path::new(&path).exists() {
            println!("## {label}: MISSING {path}");
            continue;
        }
        println!("## {label}  ({})", path.trim_start_matches(ROOT));
        println!(
            "  detect_file_format        -> {:?}",
            litchi::detect_file_format(&path)
        );
        println!(
            "  Document::open            -> {}",
            show(litchi::Document::open(&path))
        );
        println!(
            "  Presentation::open        -> {}",
            show(litchi::Presentation::open(&path))
        );
        println!(
            "  sheet::Workbook::open     -> {}",
            show(litchi::sheet::Workbook::open(&path))
        );
        println!(
            "  sheet::open_workbook      -> {}",
            show(litchi::sheet::open_workbook(&path))
        );
        println!();
    }
}
