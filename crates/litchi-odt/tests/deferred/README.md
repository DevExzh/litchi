# Deferred ODT integration coverage

The Rust sources in `legacy/` retain the default-page-layout test whose
original flat-document helper is not part of the dedicated ODT facade yet.
They remain scoped here until that flat reader exposes the semantic accessor;
active ODT tests use `Document`, `Builder`, `MutableDocument`, and the
layered style/element modules directly.
