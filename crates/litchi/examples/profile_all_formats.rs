//! Comprehensive profiling example for all supported Office file formats
//!
//! This example profiles parsing and **Markdown conversion** across all major Office formats:
//! - **Word**: DOC, DOCX, Pages
//! - **PowerPoint**: PPT, PPTX, Keynote
//! - **Excel**: XLS, XLSX, XLSB, ODS, Numbers
//!
//! Files are loaded into memory first to avoid disk I/O interference with profiling.
//! The focus is on the hot path: parsing, markdown conversion, and formatting.
//!
//! # Features
//!
//! - **Memory-based**: All files loaded into RAM to eliminate I/O from profiles
//! - **Multiple iterations**: Runs 5 iterations per file for statistical accuracy
//! - **Comprehensive metrics**: Time, throughput, markdown output size
//! - **Format coverage**: Tests all major supported Office formats
//! - **Performance focus**: Profiles the complete parsing and markdown conversion pipeline
//!
//! # Usage
//!
//! ## Basic Profiling
//!
//! ```sh
//! # Build for profiling with all features
//! cargo build --profile profiling --all-features --example profile_all_formats
//!
//! # Run with profiling (samply)
//! samply record ./target/profiling/examples/profile_all_formats
//! ```
//!
//! ## Alternative Profilers
//!
//! ```sh
//! # Using perf on Linux
//! perf record -F 99 -g ./target/profiling/examples/profile_all_formats
//! perf report
//!
//! # Using instruments on macOS
//! xcrun xctrace record --template 'Time Profiler' --output profile.trace \
//!   --launch ./target/profiling/examples/profile_all_formats
//! open profile.trace
//!
//! # Using cargo-flamegraph
//! cargo flamegraph --profile profiling --all-features \
//!   --example profile_all_formats
//! ```
//!
//! # Output
//!
//! The program generates:
//! 1. **File Loading Report**: Size and status of each test file
//! 2. **Per-Format Statistics**:
//!    - Average processing time (milliseconds)
//!    - Throughput (MB/second)
//!    - Markdown output size (characters)
//! 3. **Summary Statistics**: Overall throughput and success rate
//! 4. **Profiling Data**: For analysis in Firefox Profiler or other tools
//!
//! # Example Output
//!
//! ```text
//! ═══ Processing Documents ═══
//!   ✓ test.doc           1.8 MB │  25.57 ms │  13.81 MB/s │  35079 md chars
//!   ✓ test.pages       241.6 KB │   0.98 ms │  48.31 MB/s │   1677 md chars
//!
//! ═══ Summary ═══
//!   Files processed: 10
//!   Total iterations: 50
//!   Total time: 0.25 seconds
//!   Average throughput: 44.55 MB/s
//! ```
//!
//! # Customization
//!
//! To modify the profiling behavior:
//! - `iterations`: Number of times to process each file (default: 5, in `main()`)
//! - `profile_list.txt`: Text file listing files to profile, one per line
//!   - Supports comments (lines starting with #) and empty lines
//!   - File types are automatically detected from extensions
//!
//! # Requirements
//!
//! - Test files must be present in the workspace root
//! - `profile_list.txt` must exist in the workspace root
//! - All features must be enabled for complete format coverage
//! - Profiling build profile should be configured in Cargo.toml

use litchi::markdown::ToMarkdown;
use litchi::{Document, Presentation};
use std::collections::HashMap;
use std::time::Instant;

/// Statistics for a single file processing run
#[derive(Debug, Clone)]
struct FileStats {
    /// File name
    name: String,
    /// File size in bytes
    size: usize,
    /// Number of iterations completed
    iterations: u32,
    /// Total time spent across all iterations (microseconds)
    total_time_us: u64,
    /// Markdown output size in characters (if applicable)
    char_count: Option<usize>,
    /// Processing status
    status: ProcessStatus,
}

#[derive(Debug, Clone)]
enum ProcessStatus {
    Success,
    Failed(String),
    /// Used when certain features are disabled; currently not emitted but kept for future extensibility
    #[allow(dead_code)]
    Skipped(String),
}

impl FileStats {
    fn avg_time_ms(&self) -> f64 {
        if self.iterations == 0 {
            0.0
        } else {
            (self.total_time_us as f64) / (self.iterations as f64) / 1000.0
        }
    }

    fn throughput_mb_per_sec(&self) -> f64 {
        if self.total_time_us == 0 {
            0.0
        } else {
            let size_mb = self.size as f64 / (1024.0 * 1024.0);
            let time_sec = (self.total_time_us as f64) / 1_000_000.0;
            size_mb / time_sec
        }
    }
}

/// Profile a document file (DOC, DOCX, ODT, Pages)
fn profile_document(name: &str, data: &[u8], iterations: u32) -> FileStats {
    let mut total_time = 0u64;
    let mut char_count = None;

    for _ in 0..iterations {
        let start = Instant::now();

        match Document::from_bytes(data.to_vec()) {
            Ok(doc) => {
                // Convert to markdown - this is the hot path
                match doc.to_markdown() {
                    Ok(markdown) => {
                        char_count = Some(markdown.len());
                        // Keep the markdown in scope briefly to ensure it's not optimized away
                        std::hint::black_box(&markdown);
                    },
                    Err(e) => {
                        return FileStats {
                            name: name.to_string(),
                            size: data.len(),
                            iterations: 0,
                            total_time_us: 0,
                            char_count: None,
                            status: ProcessStatus::Failed(format!("Markdown conversion: {}", e)),
                        };
                    },
                }
            },
            Err(e) => {
                return FileStats {
                    name: name.to_string(),
                    size: data.len(),
                    iterations: 0,
                    total_time_us: 0,
                    char_count: None,
                    status: ProcessStatus::Failed(format!("Document open: {}", e)),
                };
            },
        }

        total_time += start.elapsed().as_micros() as u64;
    }

    FileStats {
        name: name.to_string(),
        size: data.len(),
        iterations,
        total_time_us: total_time,
        char_count,
        status: ProcessStatus::Success,
    }
}

/// Profile a presentation file (PPT, PPTX, ODP, Keynote)
fn profile_presentation(name: &str, data: &[u8], iterations: u32) -> FileStats {
    let mut total_time = 0u64;
    let mut char_count = None;

    for _ in 0..iterations {
        let start = Instant::now();

        match Presentation::from_bytes(data.to_vec()) {
            Ok(pres) => {
                // Convert to markdown - this is the hot path
                match pres.to_markdown() {
                    Ok(markdown) => {
                        char_count = Some(markdown.len());
                        // Keep the markdown in scope briefly to ensure it's not optimized away
                        std::hint::black_box(&markdown);
                    },
                    Err(e) => {
                        return FileStats {
                            name: name.to_string(),
                            size: data.len(),
                            iterations: 0,
                            total_time_us: 0,
                            char_count: None,
                            status: ProcessStatus::Failed(format!("Markdown conversion: {}", e)),
                        };
                    },
                }
            },
            Err(e) => {
                return FileStats {
                    name: name.to_string(),
                    size: data.len(),
                    iterations: 0,
                    total_time_us: 0,
                    char_count: None,
                    status: ProcessStatus::Failed(format!("Presentation open: {}", e)),
                };
            },
        }

        total_time += start.elapsed().as_micros() as u64;
    }

    FileStats {
        name: name.to_string(),
        size: data.len(),
        iterations,
        total_time_us: total_time,
        char_count,
        status: ProcessStatus::Success,
    }
}

/// Profile a spreadsheet file (XLS, XLSX, XLSB, ODS, Numbers)
#[cfg(any(feature = "xls", feature = "ooxml", feature = "odf", feature = "iwa"))]
fn profile_spreadsheet(name: &str, data: &[u8], iterations: u32) -> FileStats {
    use litchi::sheet::Workbook;

    let mut total_time = 0u64;
    let mut char_count = None;

    for _ in 0..iterations {
        let start = Instant::now();

        match Workbook::from_bytes(data.to_vec()) {
            Ok(workbook) => {
                // Extract text from all sheets - this is the hot path
                match workbook.text() {
                    Ok(text) => {
                        char_count = Some(text.len());
                        // Keep the text in scope briefly to ensure it's not optimized away
                        std::hint::black_box(&text);
                    },
                    Err(e) => {
                        return FileStats {
                            name: name.to_string(),
                            size: data.len(),
                            iterations: 0,
                            total_time_us: 0,
                            char_count: None,
                            status: ProcessStatus::Failed(format!("Text extraction: {}", e)),
                        };
                    },
                }
            },
            Err(e) => {
                return FileStats {
                    name: name.to_string(),
                    size: data.len(),
                    iterations: 0,
                    total_time_us: 0,
                    char_count: None,
                    status: ProcessStatus::Failed(format!("Workbook open: {}", e)),
                };
            },
        }

        total_time += start.elapsed().as_micros() as u64;
    }

    FileStats {
        name: name.to_string(),
        size: data.len(),
        iterations,
        total_time_us: total_time,
        char_count,
        status: ProcessStatus::Success,
    }
}

#[cfg(not(any(feature = "xls", feature = "ooxml", feature = "odf", feature = "iwa")))]
fn profile_spreadsheet(name: &str, data: &[u8], _iterations: u32) -> FileStats {
    FileStats {
        name: name.to_string(),
        size: data.len(),
        iterations: 0,
        total_time_us: 0,
        char_count: None,
        status: ProcessStatus::Skipped("Spreadsheet features not enabled".to_string()),
    }
}

/// Load a file into memory, returning None if it doesn't exist
fn load_file(path: &str) -> Option<Vec<u8>> {
    std::fs::read(path).ok()
}

/// Determine file type from extension
fn determine_file_type(filename: &str) -> Option<&'static str> {
    let lowercase = filename.to_lowercase();
    if lowercase.ends_with(".doc")
        || lowercase.ends_with(".docx")
        || lowercase.ends_with(".pages")
        || lowercase.ends_with(".odt")
    {
        Some("document")
    } else if lowercase.ends_with(".ppt")
        || lowercase.ends_with(".pptx")
        || lowercase.ends_with(".key")
        || lowercase.ends_with(".odp")
    {
        Some("presentation")
    } else if lowercase.ends_with(".xls")
        || lowercase.ends_with(".xlsx")
        || lowercase.ends_with(".xlsb")
        || lowercase.ends_with(".ods")
        || lowercase.ends_with(".numbers")
    {
        Some("spreadsheet")
    } else {
        None
    }
}

fn format_size(bytes: usize) -> String {
    if bytes < 1024 {
        format!("{} B", bytes)
    } else if bytes < 1024 * 1024 {
        format!("{:.1} KB", bytes as f64 / 1024.0)
    } else {
        format!("{:.1} MB", bytes as f64 / (1024.0 * 1024.0))
    }
}

fn print_stats(stats: &FileStats) {
    let status_str = match &stats.status {
        ProcessStatus::Success => "✓",
        ProcessStatus::Failed(_) => "✗",
        ProcessStatus::Skipped(_) => "⊘",
    };

    print!("  {} {:30}", status_str, stats.name);

    match &stats.status {
        ProcessStatus::Success => {
            print!(
                " {:>10} │ {:>8.2} ms │ {:>8.2} MB/s",
                format_size(stats.size),
                stats.avg_time_ms(),
                stats.throughput_mb_per_sec()
            );
            if let Some(chars) = stats.char_count {
                print!(" │ {:>8} md chars", chars);
            }
            println!();
        },
        ProcessStatus::Failed(msg) => {
            println!(" {:>10} │ FAILED: {}", format_size(stats.size), msg);
        },
        ProcessStatus::Skipped(msg) => {
            println!(" {:>10} │ SKIPPED: {}", format_size(stats.size), msg);
        },
    }
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("╔═══════════════════════════════════════════════════════════════════╗");
    println!("║  Litchi Comprehensive Format Profiling                           ║");
    println!("╚═══════════════════════════════════════════════════════════════════╝");
    println!();

    // Number of iterations per file
    let iterations = 5;

    // Read test files from profile_list.txt
    let profile_list_path = "profile_list.txt";
    let file_list_content = std::fs::read_to_string(profile_list_path).unwrap_or_else(|e| {
        eprintln!("Error reading {}: {}", profile_list_path, e);
        eprintln!("Please create a profile_list.txt file with one filename per line.");
        std::process::exit(1);
    });

    let mut files: Vec<(String, &str)> = Vec::new();
    for line in file_list_content.lines() {
        let filename = line.trim();
        if filename.is_empty() || filename.starts_with('#') {
            // Skip empty lines and comments
            continue;
        }

        if let Some(file_type) = determine_file_type(filename) {
            files.push((filename.to_string(), file_type));
        } else {
            eprintln!("Warning: Unknown file type for '{}', skipping", filename);
        }
    }

    // Load all files into memory
    println!("═══ Loading Files into Memory ═══");
    let mut file_data: HashMap<String, Vec<u8>> = HashMap::new();
    let mut total_size = 0usize;

    for (filename, _) in &files {
        if let Some(data) = load_file(filename.as_str()) {
            let size = data.len();
            total_size += size;
            println!("  ✓ {:20} {}", filename, format_size(size));
            file_data.insert(filename.clone(), data);
        } else {
            println!("  ⊘ {:20} (not found)", filename);
        }
    }

    println!();
    println!("Total data loaded: {}", format_size(total_size));
    println!("Iterations per file: {}", iterations);
    println!();

    // Process each category
    let mut all_stats = Vec::new();

    // Documents
    println!("═══ Processing Documents ═══");
    for (filename, file_type) in &files {
        if *file_type == "document"
            && let Some(data) = file_data.get(filename)
        {
            let stats = profile_document(filename.as_str(), data, iterations);
            print_stats(&stats);
            all_stats.push(stats);
        }
    }
    println!();

    // Presentations
    println!("═══ Processing Presentations ═══");
    for (filename, file_type) in &files {
        if *file_type == "presentation"
            && let Some(data) = file_data.get(filename)
        {
            let stats = profile_presentation(filename.as_str(), data, iterations);
            print_stats(&stats);
            all_stats.push(stats);
        }
    }
    println!();

    // Spreadsheets
    println!("═══ Processing Spreadsheets ═══");
    for (filename, file_type) in &files {
        if *file_type == "spreadsheet"
            && let Some(data) = file_data.get(filename)
        {
            let stats = profile_spreadsheet(filename.as_str(), data, iterations);
            print_stats(&stats);
            all_stats.push(stats);
        }
    }
    println!();

    // Summary statistics
    println!("═══ Summary ═══");
    let successful: Vec<_> = all_stats
        .iter()
        .filter(|s| matches!(s.status, ProcessStatus::Success))
        .collect();

    if !successful.is_empty() {
        let total_time_sec: f64 = successful
            .iter()
            .map(|s| s.total_time_us as f64 / 1_000_000.0)
            .sum();
        let total_iterations: u32 = successful.iter().map(|s| s.iterations).sum();
        let avg_throughput: f64 = successful
            .iter()
            .map(|s| s.throughput_mb_per_sec())
            .sum::<f64>()
            / successful.len() as f64;

        println!("  Files processed: {}", successful.len());
        println!("  Total iterations: {}", total_iterations);
        println!("  Total time: {:.2} seconds", total_time_sec);
        println!("  Average throughput: {:.2} MB/s", avg_throughput);
    }

    let failed = all_stats
        .iter()
        .filter(|s| matches!(s.status, ProcessStatus::Failed(_)))
        .count();
    if failed > 0 {
        println!("  Failed: {}", failed);
    }

    let skipped = all_stats
        .iter()
        .filter(|s| matches!(s.status, ProcessStatus::Skipped(_)))
        .count();
    if skipped > 0 {
        println!("  Skipped: {}", skipped);
    }

    println!();
    println!("Profiling complete! Profile data will be loaded in Firefox Profiler.");

    Ok(())
}
