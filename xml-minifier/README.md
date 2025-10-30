# XML Minifier

A high-performance Rust procedural macro for compile-time XML minification.

## Features

- **Compile-time minification**: XML files are minified during compilation, zero runtime overhead
- **Aggressive optimization**:
  - Removes comments and processing instructions
  - Trims and collapses whitespace in text nodes
  - Collapses empty tags (`<tag></tag>` → `<tag/>`)
  - Removes unnecessary whitespace between tags
- **Safe and standards-compliant**: Preserves XML structure and semantics
- **Fast**: Single-pass processing with efficient buffer reuse
- **Memory-efficient**: Pre-allocates buffers and uses zero-copy operations where possible

## Usage

Add to your `Cargo.toml`:

```toml
[dependencies]
xml-minifier = { path = "../xml-minifier" }
```

Use the `minified_xml!` macro:

```rust
use xml_minifier::minified_xml;

// Minify an XML file at compile time
// Path is relative to the source file calling the macro
const TEMPLATE: &str = minified_xml!("template.xml");

fn main() {
    println!("{}", TEMPLATE);
}
```

### Path Resolution

**File paths are resolved relative to the source file** that invokes the macro. This makes it intuitive to keep XML files next to your Rust source code.

#### Example Project Structure

```
my-project/
├── Cargo.toml
└── src/
    ├── main.rs
    ├── lib.rs
    └── templates/
        ├── mod.rs
        └── document.xml
```

In `src/templates/mod.rs`:
```rust
// XML file is in the same directory as this source file
const TEMPLATE: &str = minified_xml!("document.xml");
```

In `src/lib.rs`:
```rust
// XML file is in the templates subdirectory
const TEMPLATE: &str = minified_xml!("templates/document.xml");
```

## Example

Given an XML file `template.xml`:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<root>
    <!-- This is a comment -->
    <child attr="value">
        Text content
    </child>
    <empty></empty>
</root>
```

The macro produces:

```xml
<?xml version="1.0" encoding="UTF-8"?><root><child attr="value">Text content</child><empty/></root>
```

## Implementation Details

### Whitespace Handling

The minifier intelligently handles whitespace:
- Removes pure whitespace between tags
- Trims leading and trailing whitespace from text nodes
- Preserves text content

### CDATA Sections

CDATA sections are preserved as-is since they may contain formatting-sensitive content:

```xml
<root><![CDATA[Some <data> with special chars]]></root>
```

### XML Declarations

XML declarations are preserved with their attributes:

```xml
<?xml version="1.0" encoding="UTF-8"?>
```

### DOCTYPE Declarations

DOCTYPE declarations are preserved:

```xml
<!DOCTYPE html>
```

## Performance

- **Zero runtime cost**: Minification happens at compile time
- **Efficient processing**: Single-pass with buffer reuse
- **Memory-efficient**: Pre-allocates approximately half the input size
- **Zero-copy where possible**: Uses `Cow<[u8]>` and byte slices

# Tips for Rust Analyzer users

Note that the procedure macro utilizes the `local_file()` function to access the source code file,
and rust-analyzer would not correctly handle the expansion due to its limitation.
In order not to produce tons of errors and warnings, add the following settings to your VS Code settings:

```json
{
    "rust-analyzer.procMacro.ignored": {
        "xml-minifier": ["minified_xml"]
    }
}
```

## License

This is part of the Litchi project so it is licensed under the same license that the project adopts.

