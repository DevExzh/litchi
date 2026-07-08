# Building litchi-py

This document provides detailed instructions for building and developing litchi-py.

## Prerequisites

1. **Rust** (1.70 or later)
   ```bash
   curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh
   ```

2. **Python** (3.8 or later)
   ```bash
   # On macOS with Homebrew
   brew install python@3.11
   
   # On Ubuntu/Debian
   sudo apt install python3 python3-pip python3-venv
   ```

3. **Maturin** (build tool for PyO3)
   ```bash
   pip install maturin
   ```

## Development Build

For development, use `maturin develop` which builds and installs the package in the current Python environment:

```bash
cd pyo3-litchi

# Development build (faster, includes debug symbols)
maturin develop

# Release build (optimized, slower to compile)
maturin develop --release
```

After running this, you can import and use `litchi_py` in your Python code:

```python
from litchi_py import Document
doc = Document.open("test.docx")
print(doc.text())
```

## Building Wheels

To build distributable wheel files:

```bash
cd pyo3-litchi

# Build for the current platform
maturin build --release

# The wheel will be in target/wheels/
ls -lh target/wheels/
```

Install the wheel:

```bash
pip install target/wheels/litchi_py-*.whl
```

## Cross-Platform Builds

### Using maturin with Docker

Maturin can build wheels for multiple platforms using Docker:

```bash
# Install Docker first, then:

# Build for Linux (manylinux)
maturin build --release --manylinux 2014

# Build for multiple Python versions
maturin build --release --interpreter python3.8 python3.9 python3.10 python3.11 python3.12
```

### Manual Cross-Compilation

For cross-compilation, you'll need to set up Rust targets:

```bash
# For macOS (Intel)
rustup target add x86_64-apple-darwin

# For macOS (Apple Silicon)
rustup target add aarch64-apple-darwin

# For Windows
rustup target add x86_64-pc-windows-msvc

# For Linux
rustup target add x86_64-unknown-linux-gnu
```

Then build with the target:

```bash
maturin build --release --target x86_64-apple-darwin
```

## Testing

### Running Examples

```bash
cd pyo3-litchi

# Make sure the package is installed
maturin develop --release

# Run examples
python examples/document_example.py
python examples/presentation_example.py
python examples/workbook_example.py
python examples/format_detection.py
```

### Running Tests

If you have pytest installed:

```bash
pip install pytest

# Create a tests directory with your test files
mkdir -p tests
# ... add test files ...

pytest tests/
```

## Performance Profiling

To profile the Rust code:

```bash
# Build with profiling symbols
RUSTFLAGS="-C force-frame-pointers=yes" maturin develop --release

# Use your preferred profiler (e.g., py-spy for Python, perf for Linux)
pip install py-spy
py-spy record -o profile.svg -- python your_script.py
```

## Troubleshooting

### "No module named 'litchi_py'"

Make sure you've run `maturin develop` in the correct directory and your Python environment is activated.

### Compilation Errors

1. **Missing Rust**: Install Rust from https://rustup.rs
2. **Outdated Rust**: Run `rustup update`
3. **Missing dependencies**: Make sure the parent `litchi` library compiles successfully

### Linker Errors on macOS

If you get linker errors related to C++ on macOS, install Xcode Command Line Tools:

```bash
xcode-select --install
```

### Permission Errors

On Linux, if you get permission errors when installing:

```bash
# Use a virtual environment (recommended)
python3 -m venv venv
source venv/bin/activate
pip install maturin
maturin develop --release
```

## Publishing to PyPI

To publish to PyPI (for maintainers):

```bash
# Build wheels for all platforms
maturin build --release --manylinux 2014

# Upload to PyPI
maturin publish --username __token__ --password $PYPI_TOKEN
```

## IDE Setup

### VS Code

Install these extensions:
- Python (ms-python.python)
- rust-analyzer (rust-lang.rust-analyzer)
- PyO3 (ms-python.vscode-pylance for type checking)

### PyCharm

1. Enable type checking: Settings → Editor → Inspections → Python → Type Checker
2. Mark `python/` as a Sources Root

## Additional Resources

- [PyO3 User Guide](https://pyo3.rs/)
- [Maturin Documentation](https://github.com/PyO3/maturin)
- [Rust Documentation](https://doc.rust-lang.org/)

