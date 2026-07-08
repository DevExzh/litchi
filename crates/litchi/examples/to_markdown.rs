//! Comprehensive CLI tool for converting Office documents and presentations to Markdown.
//!
//! This example provides a full-featured command-line interface for converting:
//! - Word documents (.doc, .docx, .rtf)
//! - PowerPoint presentations (.ppt, .pptx)
//! - OpenDocument files (.odt, .odp)
//!
//! to Markdown format with extensive customization options.
//!
//! # Usage
//!
//! Basic conversion:
//! ```sh
//! cargo run --example to_markdown -- input.docx -o output.md
//! ```
//!
//! Convert with options:
//! ```sh
//! cargo run --example to_markdown -- input.docx -o output.md \
//!     --table-style markdown \
//!     --no-styles \
//!     --no-metadata
//! ```
//!
//! Convert to directory (auto-generates filename):
//! ```sh
//! cargo run --example to_markdown -- input.docx -o output_dir/
//! ```
//!
//! Batch conversion:
//! ```sh
//! cargo run --example to_markdown -- *.docx -o output_dir/
//! ```

use clap::{Parser, ValueEnum};
use litchi::markdown::{
    FormulaStyle, MarkdownOptions, ScriptStyle, StrikethroughStyle, TableStyle, ToMarkdown,
};
use litchi::{Document, FileFormat, Presentation, detect_file_format};
use std::fs;
use std::path::{Path, PathBuf};

/// Convert Office documents and presentations to Markdown
#[derive(Parser, Debug)]
#[command(
    name = "to_markdown",
    about = "Convert Office documents and presentations to Markdown format",
    long_about = "A high-performance CLI tool for converting Microsoft Office and OpenDocument files to Markdown.\n\
                  Supports Word documents (.doc, .docx, .rtf), PowerPoint presentations (.ppt, .pptx),\n\
                  and OpenDocument formats (.odt, .odp) with extensive customization options.",
    version
)]
struct Args {
    /// Input file(s) to convert
    #[arg(value_name = "INPUT", required = true)]
    input: Vec<PathBuf>,

    /// Output file or directory
    ///
    /// If a directory is specified (ending with /), output files will be created
    /// with the same basename as input files but with .md extension.
    /// If a single file is specified and multiple inputs are given, an error occurs.
    #[arg(short, long, value_name = "OUTPUT")]
    output: PathBuf,

    /// Table conversion style
    #[arg(long, value_enum, default_value = "markdown")]
    table_style: TableStyleArg,

    /// Formula conversion style
    #[arg(long, value_enum, default_value = "latex")]
    formula_style: FormulaStyleArg,

    /// Superscript/subscript style
    #[arg(long, value_enum, default_value = "html")]
    script_style: ScriptStyleArg,

    /// Strikethrough style
    #[arg(long, value_enum, default_value = "markdown")]
    strikethrough_style: StrikethroughStyleArg,

    /// Disable text formatting (bold, italic, etc.)
    #[arg(long)]
    no_styles: bool,

    /// Disable document metadata (YAML front matter)
    #[arg(long)]
    no_metadata: bool,

    /// Force overwrite existing files
    #[arg(short, long)]
    force: bool,

    /// Verbose output
    #[arg(short, long)]
    verbose: bool,
}

/// Table style options for CLI
#[derive(Debug, Clone, Copy, ValueEnum)]
enum TableStyleArg {
    /// Standard Markdown tables
    Markdown,
    /// Minimal HTML tables
    MinimalHtml,
    /// Styled HTML tables with CSS classes
    StyledHtml,
}

impl From<TableStyleArg> for TableStyle {
    fn from(arg: TableStyleArg) -> Self {
        match arg {
            TableStyleArg::Markdown => TableStyle::Markdown,
            TableStyleArg::MinimalHtml => TableStyle::MinimalHtml,
            TableStyleArg::StyledHtml => TableStyle::StyledHtml,
        }
    }
}

/// Formula style options for CLI
#[derive(Debug, Clone, Copy, ValueEnum)]
enum FormulaStyleArg {
    /// LaTeX math mode (\( \) and \[ \])
    Latex,
    /// Dollar signs ($ and $$) - GitHub flavored
    Dollar,
}

impl From<FormulaStyleArg> for FormulaStyle {
    fn from(arg: FormulaStyleArg) -> Self {
        match arg {
            FormulaStyleArg::Latex => FormulaStyle::LaTeX,
            FormulaStyleArg::Dollar => FormulaStyle::Dollar,
        }
    }
}

/// Script (superscript/subscript) style options for CLI
#[derive(Debug, Clone, Copy, ValueEnum)]
enum ScriptStyleArg {
    /// HTML tags (<sup>, <sub>)
    Html,
    /// Unicode characters
    Unicode,
}

impl From<ScriptStyleArg> for ScriptStyle {
    fn from(arg: ScriptStyleArg) -> Self {
        match arg {
            ScriptStyleArg::Html => ScriptStyle::Html,
            ScriptStyleArg::Unicode => ScriptStyle::Unicode,
        }
    }
}

/// Strikethrough style options for CLI
#[derive(Debug, Clone, Copy, ValueEnum)]
enum StrikethroughStyleArg {
    /// HTML <del> tag
    Html,
    /// Markdown ~~text~~
    Markdown,
}

impl From<StrikethroughStyleArg> for StrikethroughStyle {
    fn from(arg: StrikethroughStyleArg) -> Self {
        match arg {
            StrikethroughStyleArg::Html => StrikethroughStyle::Html,
            StrikethroughStyleArg::Markdown => StrikethroughStyle::Markdown,
        }
    }
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args = Args::parse();

    // Validate inputs
    for input in &args.input {
        if !input.exists() {
            eprintln!("Error: Input file does not exist: {}", input.display());
            std::process::exit(1);
        }
        if !input.is_file() {
            eprintln!("Error: Input path is not a file: {}", input.display());
            std::process::exit(1);
        }
    }

    // Determine if output is a directory
    let output_is_dir = args.output.to_string_lossy().ends_with('/')
        || args.output.to_string_lossy().ends_with('\\')
        || (args.output.exists() && args.output.is_dir());

    // Validate output
    if args.input.len() > 1 && !output_is_dir {
        eprintln!("Error: Multiple input files require output to be a directory");
        std::process::exit(1);
    }

    // Create output directory if needed
    if output_is_dir && !args.output.exists() {
        if args.verbose {
            println!("Creating output directory: {}", args.output.display());
        }
        fs::create_dir_all(&args.output)?;
    }

    // Build markdown options
    let options = MarkdownOptions::new()
        .with_table_style(args.table_style.into())
        .with_formula_style(args.formula_style.into())
        .with_script_style(args.script_style.into())
        .with_strikethrough_style(args.strikethrough_style.into())
        .with_styles(!args.no_styles)
        .with_metadata(!args.no_metadata);

    // Process each input file
    let mut success_count = 0;
    let mut error_count = 0;

    for input in &args.input {
        let output_path = if output_is_dir {
            // Generate output filename from input
            let stem = input
                .file_stem()
                .ok_or_else(|| format!("Invalid input filename: {}", input.display()))?
                .to_string_lossy();
            args.output.join(format!("{}.md", stem))
        } else {
            args.output.clone()
        };

        // Check if output exists and we're not forcing overwrite
        if output_path.exists() && !args.force {
            eprintln!(
                "Error: Output file already exists: {}",
                output_path.display()
            );
            eprintln!("       Use --force to overwrite");
            error_count += 1;
            continue;
        }

        if args.verbose {
            println!(
                "Converting: {} -> {}",
                input.display(),
                output_path.display()
            );
        }

        match convert_file(input, &output_path, &options, args.verbose) {
            Ok(()) => {
                success_count += 1;
                if !args.verbose {
                    println!("✓ {}", input.display());
                }
            },
            Err(e) => {
                error_count += 1;
                eprintln!("✗ {}: {}", input.display(), e);
            },
        }
    }

    // Print summary
    if args.input.len() > 1 {
        println!("\n=== Conversion Summary ===");
        println!("Success: {}", success_count);
        println!("Failed:  {}", error_count);
        println!("Total:   {}", args.input.len());
    }

    if error_count > 0 {
        std::process::exit(1);
    }

    Ok(())
}

/// Convert a single file to Markdown
fn convert_file(
    input: &Path,
    output: &Path,
    options: &MarkdownOptions,
    verbose: bool,
) -> Result<(), Box<dyn std::error::Error>> {
    // Detect file format
    let format = detect_file_format(input)
        .ok_or_else(|| format!("Could not detect file format for: {}", input.display()))?;

    if verbose {
        println!("  Detected format: {:?}", format);
    }

    // Convert based on format
    let markdown = match format {
        FileFormat::Doc | FileFormat::Docx | FileFormat::Odt | FileFormat::Rtf => {
            // Word document formats
            let doc = Document::open(input)?;

            if verbose {
                println!("  Processing document...");
            }

            doc.to_markdown_with_options(options)?
        },
        FileFormat::Ppt | FileFormat::Pptx | FileFormat::Odp => {
            // PowerPoint presentation formats
            let pres = Presentation::open(input)?;

            if verbose {
                let slide_count = pres.slide_count()?;
                println!("  Processing presentation ({} slides)...", slide_count);
            }

            pres.to_markdown_with_options(options)?
        },
        _ => {
            return Err(format!(
                "Unsupported file format: {:?}. Only documents and presentations are supported.",
                format
            )
            .into());
        },
    };

    // Write output
    if verbose {
        println!("  Writing output ({} bytes)...", markdown.len());
    }

    fs::write(output, markdown)?;

    if verbose {
        println!("  ✓ Conversion complete");
    }

    Ok(())
}
