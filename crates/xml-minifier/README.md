# XML Minifier

A bounded XML compactness auditor with compile-time producer-template macros.

## Features

- **Compile-time compactness enforcement**: producer templates are checked and normalized during compilation with zero runtime overhead
- **Semantic preservation**:
  - Removes authoring comments but preserves processing instructions
  - Rejects CR/LF/tab-only structural text outside inherited `xml:space="preserve"`
  - Preserves every accepted text event byte-for-byte, including pure-space content
  - Collapses empty tags (`<tag></tag>` → `<tag/>`)
- **Source contract**: producer templates must contain no non-semantic formatting whitespace; violations produce `compile_error!`
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
<?xml version="1.0" encoding="UTF-8"?><root><!-- This is a comment --><child attr="value">Text content</child><empty></empty></root>
```

The macro produces:

```xml
<?xml version="1.0" encoding="UTF-8"?><root><child attr="value">Text content</child><empty/></root>
```

## Implementation Details

### Whitespace Handling

The macros never infer an XML schema content model. They reject text consisting
only of XML whitespace when it contains CR, LF, or tab outside inherited
`xml:space="preserve"`; they never silently delete it. Pure-space nodes and all
text containing semantic characters are preserved byte-for-byte, including any
CR, LF, or tab within that semantic text. Keep producer-template sources compact
and use `audit::verify` to enforce the same structural whitespace contract.

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
- **Memory-efficient**: Pre-allocates at most the bounded input size
- **Zero-copy where possible**: Uses borrowed parser events and byte slices

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
## Compact output auditing

`xml-minifier` also exposes a bounded, non-rewriting verifier for generated
OOXML and ODF parts and checked-in XML assets:

```rust
use xml_minifier::audit::{self, Limits};

let report = audit::verify(b"<root key=\"value\">text</root>", Limits::default())?;
assert_eq!(report.max_depth(), 1);
# Ok::<(), audit::Error>(())
```

The verifier rejects structural CR/LF/tab indentation, whitespace before a tag
close, and nonminimal attribute separators. Character data, CDATA, and plain
space-only nodes are never normalized. An inherited `xml:space="preserve"`
keeps line-oriented whitespace exact. `audit::package` verifies borrowed named
parts under aggregate part and byte budgets without rewriting opaque input.
DTD and DOCTYPE declarations are rejected without resolving entities.
