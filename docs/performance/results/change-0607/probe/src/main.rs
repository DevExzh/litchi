//! Change 0607 probe: reachability of the PPTX eager-save slide regeneration,
//! and the cost of the route that is reachable.
//!
//! Subcommands:
//!   admission <dir>            census: can an opened deck reach the mutable writer?
//!   census <slides>            deterministic counts for one authored edit-and-save
//!   resave <slides> <samples>  build once, then `samples` x (edit one title + to_bytes)
//!   nopsave <slides> <samples> build once, then `samples` x (to_bytes with a clean model)
//!   create <slides> <samples>  `samples` x (author the whole deck + to_bytes)

use std::env;
use std::error::Error;
use std::fs;
use std::path::{Path, PathBuf};
use std::time::Instant;

use litchi_pptx::Package;

const TITLE_WIDTH: usize = 24;

fn body_text(slide: usize, box_index: usize) -> String {
    format!(
        "Slide {slide:04} line {box_index} - a sentence of ordinary presentation body text that a real deck would carry."
    )
}

fn fixed_title(slide: usize, revision: u64) -> String {
    let title = format!("S{slide:04} rev {revision:06}");
    debug_assert!(title.len() <= TITLE_WIDTH);
    format!("{title:<TITLE_WIDTH$}")
}

fn build_deck(slides: usize) -> Result<Package, Box<dyn Error + Send + Sync>> {
    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        for index in 0..slides {
            let slide = presentation.add_slide()?;
            slide.set_title(&fixed_title(index, 0));
            for box_index in 0..3 {
                slide.add_text_box(
                    &body_text(index, box_index),
                    914_400,
                    1_828_800 + 914_400 * i64::try_from(box_index)?,
                    7_315_200,
                    914_400,
                );
            }
            if index % 4 == 3 {
                slide.set_notes(&format!("Speaker notes for slide {index:04}."));
            }
        }
    }
    Ok(package)
}

fn collect_pptx(dir: &Path, out: &mut Vec<PathBuf>) -> std::io::Result<()> {
    for entry in fs::read_dir(dir)? {
        let entry = entry?;
        let path = entry.path();
        if path.is_dir() {
            collect_pptx(&path, out)?;
        } else if path.extension().is_some_and(|ext| ext == "pptx") {
            out.push(path);
        }
    }
    Ok(())
}

fn admission(dir: &str) -> Result<(), Box<dyn Error + Send + Sync>> {
    let mut files = Vec::new();
    collect_pptx(Path::new(dir), &mut files)?;
    files.sort();
    let (mut opened, mut open_refused, mut writer_admitted, mut writer_refused) = (0, 0, 0, 0);
    println!("bytes\topen\tpresentation_mut\tis_modified\tpath\tdetail");
    for path in &files {
        let bytes = fs::metadata(path).map(|meta| meta.len()).unwrap_or(0);
        match Package::open(path) {
            Err(error) => {
                open_refused += 1;
                println!("{bytes}\tERR\t-\t-\t{}\t{error}", path.display());
            },
            Ok(mut package) => {
                opened += 1;
                let modified = package.is_modified();
                match package.presentation_mut() {
                    Ok(_) => {
                        writer_admitted += 1;
                        println!("{bytes}\tok\tADMITTED\t{modified}\t{}\t-", path.display());
                    },
                    Err(error) => {
                        writer_refused += 1;
                        println!(
                            "{bytes}\tok\trefused\t{modified}\t{}\t{error}",
                            path.display()
                        );
                    },
                }
            },
        }
    }
    println!(
        "# files={} opened={opened} open_refused={open_refused} writer_admitted={writer_admitted} writer_refused={writer_refused}",
        files.len()
    );
    Ok(())
}

fn census(slides: usize) -> Result<(), Box<dyn Error + Send + Sync>> {
    let mut package = build_deck(slides)?;
    let first = package.to_bytes()?;
    // One edit through the mutable model.
    {
        let presentation = package.presentation_mut()?;
        let slide = presentation
            .slide_mut(0)
            .ok_or("authored deck has no slide 0")?;
        slide.set_title(&fixed_title(0, 1));
    }
    let modified: Vec<usize> = {
        let presentation = package.presentation_mut()?;
        (0..presentation.slide_count())
            .filter(|index| {
                presentation
                    .slides()
                    .get(*index)
                    .is_some_and(litchi_pptx::writer::MutableSlide::is_modified)
            })
            .collect()
    };
    let (mut generated_bytes, mut notes) = (0usize, 0usize);
    {
        let presentation = package.presentation_mut()?;
        for slide in presentation.slides() {
            generated_bytes += slide.generate_slide_xml()?.len();
            if slide.has_notes() {
                notes += 1;
            }
        }
    }
    let second = package.to_bytes()?;
    println!("slides\t{slides}");
    println!("notes_parts\t{notes}");
    println!("slides_modified\t{}", modified.len());
    println!("slides_unmodified\t{}", slides - modified.len());
    println!("slide_parts_removed_and_regenerated_per_save\t{slides}");
    println!("notes_parts_removed_and_regenerated_per_save\t{notes}");
    println!("slide_xml_bytes_generated_per_save\t{generated_bytes}");
    println!("first_save_bytes\t{}", first.len());
    println!("second_save_bytes\t{}", second.len());
    Ok(())
}

fn resave(slides: usize, samples: u64) -> Result<(), Box<dyn Error + Send + Sync>> {
    let mut package = build_deck(slides)?;
    let mut total = 0usize;
    let _ = package.to_bytes()?;
    let start = Instant::now();
    for revision in 1..=samples {
        {
            let presentation = package.presentation_mut()?;
            let slide = presentation
                .slide_mut(0)
                .ok_or("authored deck has no slide 0")?;
            slide.set_title(&fixed_title(0, revision));
        }
        total += package.to_bytes()?.len();
    }
    let elapsed = start.elapsed();
    println!("{}\t{}\t{total}", elapsed.as_nanos(), samples);
    Ok(())
}

fn nopsave(slides: usize, samples: u64) -> Result<(), Box<dyn Error + Send + Sync>> {
    let mut package = build_deck(slides)?;
    let mut total = 0usize;
    let _ = package.to_bytes()?;
    let start = Instant::now();
    for _ in 0..samples {
        total += package.to_bytes()?.len();
    }
    let elapsed = start.elapsed();
    println!("{}\t{}\t{total}", elapsed.as_nanos(), samples);
    Ok(())
}

fn create(slides: usize, samples: u64) -> Result<(), Box<dyn Error + Send + Sync>> {
    let mut total = 0usize;
    let start = Instant::now();
    for _ in 0..samples {
        let mut package = build_deck(slides)?;
        total += package.to_bytes()?.len();
    }
    let elapsed = start.elapsed();
    println!("{}\t{}\t{total}", elapsed.as_nanos(), samples);
    Ok(())
}


fn stability(slides: usize, out: &str) -> Result<(), Box<dyn Error + Send + Sync>> {
    stability_with(slides, out, 0, "save1.pptx", "save2.pptx")
}

fn editdiff(slides: usize, out: &str) -> Result<(), Box<dyn Error + Send + Sync>> {
    stability_with(slides, out, 1, "edit-before.pptx", "edit-after.pptx")
}

fn stability_with(
    slides: usize,
    out: &str,
    revision: u64,
    first_name: &str,
    second_name: &str,
) -> Result<(), Box<dyn Error + Send + Sync>> {
    let mut package = build_deck(slides)?;
    let first = package.to_bytes()?;
    // Re-assert the identical title: the model is marked modified, but every
    // slide's value is unchanged, so a value-identical writer must reproduce
    // the first save byte for byte.
    {
        let presentation = package.presentation_mut()?;
        let slide = presentation
            .slide_mut(0)
            .ok_or("authored deck has no slide 0")?;
        slide.set_title(&fixed_title(0, revision));
    }
    let second = package.to_bytes()?;
    fs::write(format!("{out}/{first_name}"), &first)?;
    fs::write(format!("{out}/{second_name}"), &second)?;
    println!("first_bytes\t{}", first.len());
    println!("second_bytes\t{}", second.len());
    println!("identical\t{}", first == second);
    if first != second {
        let common = first.len().min(second.len());
        let at = (0..common).find(|i| first[*i] != second[*i]);
        println!("first_difference_at\t{at:?}");
    }
    Ok(())
}


fn opened(dir: &str) -> Result<(), Box<dyn Error + Send + Sync>> {
    let mut files = Vec::new();
    collect_pptx(Path::new(dir), &mut files)?;
    files.sort();
    let (mut exact, mut differing, mut refused) = (0, 0, 0);
    println!("source_bytes\tsaved_bytes\texact\tpath\tdetail");
    for path in &files {
        let source = fs::read(path)?;
        let mut package = match Package::from_vec(source.clone()) {
            Ok(package) => package,
            Err(error) => {
                refused += 1;
                println!("{}\t-\tERR\t{}\t{error}", source.len(), path.display());
                continue;
            },
        };
        match package.to_bytes() {
            Ok(saved) => {
                let same = saved == source;
                if same { exact += 1 } else { differing += 1 }
                println!(
                    "{}\t{}\t{same}\t{}\t-",
                    source.len(),
                    saved.len(),
                    path.display()
                );
            },
            Err(error) => {
                refused += 1;
                println!("{}\t-\tERR\t{}\t{error}", source.len(), path.display());
            },
        }
    }
    println!(
        "# files={} exact={exact} differing={differing} refused={refused}",
        files.len()
    );
    Ok(())
}

fn main() -> Result<(), Box<dyn Error + Send + Sync>> {
    let args: Vec<String> = env::args().collect();
    match args.get(1).map(String::as_str) {
        Some("admission") => admission(args.get(2).map_or("test-data", String::as_str)),
        Some("census") => census(args[2].parse()?),
        Some("opened") => opened(args.get(2).map_or("test-data", String::as_str)),
        Some("stability") => stability(args[2].parse()?, &args[3]),
        Some("editdiff") => editdiff(args[2].parse()?, &args[3]),
        Some("resave") => resave(args[2].parse()?, args[3].parse()?),
        Some("nopsave") => nopsave(args[2].parse()?, args[3].parse()?),
        Some("create") => create(args[2].parse()?, args[3].parse()?),
        _ => {
            eprintln!(
                "usage: probe0607 admission <dir> | census <slides> | stability <slides> <outdir> | resave <slides> <samples> | nopsave <slides> <samples> | create <slides> <samples>"
            );
            Ok(())
        },
    }
}
