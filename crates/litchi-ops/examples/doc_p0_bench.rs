//! Paired DOC P0 timing (record 0258): the primary native package
//! transaction (`litchi_doc::Package::edit`) and native reader against the
//! unified `File` calls that wrap them.
//!
//! Each generated document has `N` body paragraphs of two runs (`plain` and
//! bold); whole-paragraph text replacement uses a sibling document of `N`
//! single-run paragraphs. Targets are paragraph `N / 2`. Owners, parsed paths and properties
//! are prepared outside timing; every timed sample starts from the same
//! immutable owner. Output bytes of both legs are compared outside timing.
//! Two warm-up rounds, then ABBA rounds (native, unified, unified, native).
#[cfg(feature = "doc")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use litchi_core::Position;
    use litchi_core::selector::Path;
    use litchi_doc::body_text::{CharacterProperty, CharacterValue, TextTarget};
    use litchi_doc::writer::{CharacterFormatting, ParagraphFormatting};
    use litchi_ops::{NewNode, Placement, Props, open_bytes};
    use std::hint::black_box;
    use std::time::Instant;

    type Leg<'a> = Box<dyn Fn() -> Result<Option<Vec<u8>>, Box<dyn std::error::Error>> + 'a>;

    fn median(values: &mut [u128]) -> u128 {
        values.sort_unstable();
        values[values.len() / 2]
    }

    let mut args = std::env::args().skip(1);
    let mut rounds = 9usize;
    let mut counts = vec![1usize, 10, 100, 1_000];
    while let Some(arg) = args.next() {
        match arg.as_str() {
            "--rounds" => rounds = args.next().ok_or("--rounds needs a value")?.parse()?,
            "--counts" => {
                counts = args
                    .next()
                    .ok_or("--counts needs a value")?
                    .split(',')
                    .map(str::parse)
                    .collect::<Result<_, _>>()?;
            },
            other => return Err(format!("unknown argument {other}").into()),
        }
    }

    println!("[");
    let mut first = true;
    for count in counts {
        let mut builder = litchi_doc::Builder::new();
        for index in 0..count {
            builder.add_paragraph_runs(
                vec![
                    (format!("plain {index:05} "), CharacterFormatting::default()),
                    (
                        format!("bold {index:05}"),
                        CharacterFormatting {
                            bold: Some(true),
                            ..CharacterFormatting::default()
                        },
                    ),
                ],
                ParagraphFormatting::default(),
            )?;
        }
        let bytes = builder.build()?.to_bytes()?;
        let package = litchi_doc::Package::from_bytes(&bytes)?;
        let file = open_bytes(bytes.clone())?;
        // Whole-paragraph text replacement needs one uniform run.
        let mut builder = litchi_doc::Builder::new();
        for index in 0..count {
            builder.add_paragraph(&format!("seed paragraph {index:05}"))?;
        }
        let plain_bytes = builder.build()?.to_bytes()?;
        let plain_package = litchi_doc::Package::from_bytes(&plain_bytes)?;
        let plain_file = open_bytes(plain_bytes)?;
        let target = count / 2;
        let position = Position::new(target);
        let paragraph_path = Path::parse(&format!("/body/p[{}]", target + 1))?.into_owned();
        let run_path = Path::parse(&format!("/body/p[{}]/r[2]", target + 1))?.into_owned();
        let body_path = Path::parse("/body")?.into_owned();
        let run_span = TextTarget::BodySpan {
            paragraph: position,
            start: 12,
            end: 22,
        };
        let replacement = format!("updated paragraph {target:05} with a longer payload");

        let scenarios: Vec<(&str, Leg<'_>, Leg<'_>)> = vec![
            (
                "paragraph_text_set",
                Box::new(|| {
                    let mut edit = plain_package.edit()?;
                    edit.replace_paragraph(position, &replacement)?;
                    Ok(Some(black_box(edit.commit()?.into_package()).to_bytes()?))
                }),
                Box::new(|| {
                    let props = Props::new().with("text", replacement.as_str());
                    let output = black_box(plain_file.set(&paragraph_path, props)?);
                    Ok(Some(output.to_bytes()?))
                }),
            ),
            (
                "run_format_set",
                Box::new(|| {
                    let mut edit = package.edit()?;
                    edit.set_character_property(run_span, CharacterProperty::Bold, false)?;
                    edit.set_character_value(run_span, CharacterValue::Size(28))?;
                    Ok(Some(black_box(edit.commit()?.into_package()).to_bytes()?))
                }),
                Box::new(|| {
                    let props = Props::new().with("bold", false).with("size", "14pt");
                    let output = black_box(file.set(&run_path, props)?);
                    Ok(Some(output.to_bytes()?))
                }),
            ),
            (
                "paragraph_insert",
                Box::new(|| {
                    let mut edit = package.edit()?;
                    edit.insert_paragraph(position, "inserted")?;
                    Ok(Some(black_box(edit.commit()?.into_package()).to_bytes()?))
                }),
                Box::new(|| {
                    let node = NewNode::new("p")
                        .prop("text", "inserted")
                        .at(Placement::Position(position));
                    let output = black_box(file.add(&body_path, node)?);
                    Ok(Some(output.file.to_bytes()?))
                }),
            ),
            (
                "paragraph_text_get",
                Box::new(|| {
                    let document = package.document();
                    let text = document.body_paragraph_text(position)?.map(str::to_owned);
                    black_box(text);
                    Ok(None)
                }),
                Box::new(|| {
                    let node = file.get(&paragraph_path)?;
                    let text = node.view().text()?.map(std::borrow::Cow::into_owned);
                    black_box(text);
                    Ok(None)
                }),
            ),
            (
                "run_bold_get",
                Box::new(|| {
                    let document = package.document();
                    let paragraph = document.body_paragraph(position)?.ok_or("paragraph")?;
                    black_box(paragraph.runs()?[1].bold());
                    Ok(None)
                }),
                Box::new(|| {
                    let node = file.get(&run_path)?;
                    let bold = node.view().format()?.find(|prop| prop.key() == "bold");
                    black_box(bold.map(|prop| prop.to_string()));
                    Ok(None)
                }),
            ),
        ];

        for (name, native, unified) in &scenarios {
            // Exactness outside timing: both legs publish identical bytes.
            if native()? != unified()? {
                return Err(format!("{name} N={count}: native and unified outputs differ").into());
            }
            let mut native_ns = Vec::new();
            let mut unified_ns = Vec::new();
            for round in 0..rounds + 2 {
                for route in [false, true, true, false] {
                    let start = Instant::now();
                    let output = if route { unified()? } else { native()? };
                    let elapsed = start.elapsed().as_nanos();
                    drop(black_box(output));
                    if round >= 2 {
                        if route {
                            unified_ns.push(elapsed);
                        } else {
                            native_ns.push(elapsed);
                        }
                    }
                }
            }
            let native_median = median(&mut native_ns);
            let unified_median = median(&mut unified_ns);
            if !first {
                println!(",");
            }
            first = false;
            print!(
                "  {{\"scenario\": \"{name}\", \"paragraphs\": {count}, \"samples_per_leg\": {}, \"native_median_ns\": {native_median}, \"unified_median_ns\": {unified_median}, \"ratio\": {:.4}, \"native_ns\": {native_ns:?}, \"unified_ns\": {unified_ns:?}}}",
                native_ns.len(),
                unified_median as f64 / native_median as f64
            );
        }
    }
    println!("\n]");
    Ok(())
}

#[cfg(not(feature = "doc"))]
fn main() {
    eprintln!("enable the `doc` feature");
}
