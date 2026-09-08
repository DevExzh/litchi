# Keynote fuzz targets

The `keynote_slide_audio_creation` target keeps the checked-in native
`media-comments-baseline-native.key` package as its source fixture and uses
the small `target-fresh-audio` corpus input to reach a successful creation
transaction. It also mutates bounded selectors, filenames, placement,
duration, and WAV payloads while checking exact source preservation, replay,
inverse restoration, and double-inverse replay.

Run a bounded smoke pass from this directory with:

```sh
cargo +nightly fuzz run keynote_slide_audio_creation -- -runs=256 -max_len=4096 -timeout=60 -rss_limit_mb=2048
```

The default fuzz build uses AddressSanitizer. Successful seed creation is
mandatory, so a restrictive budget cannot turn the entire campaign into
rejection-only coverage. Use a scratch corpus directory when preserving the
checked-in seeds unchanged; remove generated artifacts and the standalone
build directory after the campaign.

The `keynote_slide_movie_creation` target uses the same native package and
borrows its source-order file movie's video and poster bytes. Its checked-in
`target-file-movie` seed reaches a successful transaction; mutations cover
both asset payloads, both filenames, selectors, finite placement/dimensions,
duration, replay, inverse restoration, and bounded semantic rejection.

Run its bounded smoke pass with:

```sh
cargo +nightly fuzz run keynote_slide_movie_creation -- -runs=256 -max_len=4096 -timeout=60 -rss_limit_mb=2048
```

The `keynote_slide_movie_geometry` target uses the fresh native movie fixture
to exercise all-media source-order selection, composed geometry and transform
edits, original-size restoration, exact replay, and inverse restoration.  Run
its focused smoke pass with:

```sh
cargo +nightly fuzz run keynote_slide_movie_geometry -- -runs=256 -max_len=4096 -timeout=60 -rss_limit_mb=2048
```

The `keynote_slide_media_properties` target also reads the permanent native
placeholder fixture, verifies its source-order `Placeholder` classification
and Unicode accessibility marker, and checks that placeholder/live-video
property mutations are rejected without changing the source bytes.  Its
bounded smoke pass is:

```sh
cargo +nightly fuzz run keynote_slide_media_properties -- -runs=256 -max_len=4096 -timeout=60 -rss_limit_mb=2048
```
