# litchi-cfb

Parser and writer for the Microsoft Compound File Binary (CFB / OLE2) container format.

## Overview

CFB is the storage substrate underneath the legacy Microsoft Office binary
documents (`.doc`, `.xls`, `.ppt`) and is also used to wrap encrypted OOXML
packages. This crate implements `[MS-CFB]` reading and writing and exposes the
directory tree and stream I/O. Shared OLE Property Set grammar is owned by
`litchi-ole-common`, which consumes this CFB substrate. The container is
consumed directly by `litchi-doc`, `litchi-ppt`, `litchi-xls`, umbrella format
detection, and OOXML encryption.

## Usage

```toml
[dependencies]
litchi-cfb = "0.0.1"
```

```rust
use litchi_cfb::OleFile;
use std::fs::File;

let file = File::open("example.doc")?;
let mut ole = OleFile::open(file)?;
let word_doc = ole.open_stream(&["WordDocument"])?;
# Ok::<(), litchi_cfb::OleError>(())
```

## Features

- Zero-copy parsing of CFB headers, FAT/MiniFAT, and directory entries
- Stream extraction by path through the storage tree
- CFB storage primitives consumed by the shared OLE Property Set owner
- Optional `write` feature for authoring new CFB containers via `OleWriter`

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
