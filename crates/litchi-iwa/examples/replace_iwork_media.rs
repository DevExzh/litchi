use std::env;
use std::fs;
use std::path::Path;

use litchi_iwa::IWorkMediaEditor;
use litchi_iwa::keynote::KeynoteEditor;
use litchi_iwa::media::MediaAssetId;
use litchi_iwa::numbers::NumbersEditor;
use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let arguments = env::args().skip(1).collect::<Vec<_>>();
    let Some(input) = arguments.first() else {
        eprintln!(
            "usage: replace_iwork_media <input> [<output> (<data-identifier> <replacement-file> | insert <media-file> | remove <data-identifier>)]"
        );
        std::process::exit(2);
    };

    let editor = IWorkMediaEditor::open(input)?;
    for asset in editor.assets() {
        println!(
            "id={} type={} path={} actual={:?} declared={:?} component_records={} component_refs={} message_refs={} metadata={}",
            asset.data_identifier,
            asset.media_type.name(),
            asset.package_path.as_deref().unwrap_or("<unmaterialized>"),
            asset.size,
            asset.declared_size,
            asset.component_reference_record_count,
            asset.component_reference_count,
            asset.message_reference_count,
            asset.has_data_metadata,
        );
    }

    if arguments.len() == 4 {
        let output = &arguments[1];
        if arguments[2] == "insert" {
            let data = fs::read(&arguments[3])?;
            let preferred_filename = Path::new(&arguments[3])
                .file_name()
                .and_then(|name| name.to_str())
                .ok_or("media path has no UTF-8 basename")?;
            let mut media = IWorkMediaEditor::open(input)?;
            let inserted = media.insert_unreferenced(preferred_filename, &data)?;
            media.save(output)?;
            println!(
                "inserted id={} path={}; saved {output}",
                inserted.data_identifier,
                inserted
                    .package_path
                    .as_deref()
                    .unwrap_or("<unmaterialized>")
            );
            return Ok(());
        }
        if arguments[2] == "remove" {
            let data_identifier = MediaAssetId::try_from(arguments[3].parse::<u64>()?)?;
            let mut media = IWorkMediaEditor::open(input)?;
            let removed = media.remove_unreferenced(data_identifier)?;
            media.save(output)?;
            println!(
                "removed id={data_identifier} bytes={:?}; saved {output}",
                removed.as_ref().map(Vec::len)
            );
            return Ok(());
        }

        // The application-specific wrapper methods still use their native
        // wire ID. Keep that value local while the shared media editor uses
        // the checked semantic identifier above it.
        let raw_data_identifier = arguments[2].parse::<u64>()?;
        let data_identifier = MediaAssetId::try_from(raw_data_identifier)?;
        let replacement = fs::read(&arguments[3])?;
        let previous = match Path::new(input)
            .extension()
            .and_then(|extension| extension.to_str())
        {
            Some("key") => {
                let mut app = KeynoteEditor::open(input)?;
                let previous = app.replace_media(raw_data_identifier, &replacement)?;
                app.save(output)?;
                previous
            },
            Some("numbers") => {
                let mut app = NumbersEditor::open(input)?;
                let previous = app.replace_media(raw_data_identifier, &replacement)?;
                app.save(output)?;
                previous
            },
            Some("pages") => {
                let mut app = PagesEditor::open(input)?;
                let previous = app.replace_media(raw_data_identifier, &replacement)?;
                app.save(output)?;
                previous
            },
            extension => return Err(format!("unsupported iWork extension: {extension:?}").into()),
        };
        println!(
            "replaced id={data_identifier}: {} -> {} bytes; saved {output}",
            previous.len(),
            replacement.len()
        );
    } else if arguments.len() != 1 {
        eprintln!(
            "usage: replace_iwork_media <input> [<output> (<data-identifier> <replacement-file> | insert <media-file> | remove <data-identifier>)]"
        );
        std::process::exit(2);
    }

    Ok(())
}
