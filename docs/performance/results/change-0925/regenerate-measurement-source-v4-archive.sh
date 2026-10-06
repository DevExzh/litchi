#!/bin/sh
set -eu
source_root=${1:?usage: $0 SOURCE_ROOT ARCHIVE_PATH}
archive_path=${2:?usage: $0 SOURCE_ROOT ARCHIVE_PATH}
exec tar --format=gnu --sort=name --mtime='UTC 1970-01-01 00:00:00' --owner=0 --group=0 --numeric-owner -cf "$archive_path" -C "$source_root" .
