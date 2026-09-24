#!/usr/bin/env bash
set -euo pipefail

if [[ $# -ne 1 ]]; then
    echo "usage: host_probe.sh OUTPUT" >&2
    exit 2
fi

OUTPUT=$1
if [[ "$OUTPUT" == "/" || -e "$OUTPUT" || -L "$OUTPUT" ]]; then
    echo "host probe output must be a fresh file" >&2
    exit 2
fi
mkdir -p -- "$(dirname -- "$OUTPUT")"
umask 077

cpu_model=unavailable
logical_cpus=unavailable
if [[ -r /proc/cpuinfo ]]; then
    cpu_model=$(awk -F: '/^(model name|Hardware)[[:space:]]*:/ {value=$2; sub(/^[[:space:]]+/, "", value); print value; exit}' /proc/cpuinfo || true)
    logical_cpus=$(awk '/^processor[[:space:]]*:/ {count++} END {print count + 0}' /proc/cpuinfo || true)
fi
[[ -n "$cpu_model" ]] || cpu_model=unavailable
[[ -n "$logical_cpus" ]] || logical_cpus=unavailable

memory_total_kib=unavailable
memory_available_kib=unavailable
if [[ -r /proc/meminfo ]]; then
    memory_total_kib=$(awk '$1 == "MemTotal:" {print $2; exit}' /proc/meminfo || true)
    memory_available_kib=$(awk '$1 == "MemAvailable:" {print $2; exit}' /proc/meminfo || true)
fi
[[ -n "$memory_total_kib" ]] || memory_total_kib=unavailable
[[ -n "$memory_available_kib" ]] || memory_available_kib=unavailable

load_1m=unavailable
if [[ -r /proc/loadavg ]]; then
    load_1m=$(awk '{print $1; exit}' /proc/loadavg || true)
fi
[[ -n "$load_1m" ]] || load_1m=unavailable

os_name=unavailable
if [[ -r /etc/os-release ]]; then
    os_name=$(awk -F= '$1 == "PRETTY_NAME" {value=$2; gsub(/^"|"$/, "", value); print value; exit}' /etc/os-release || true)
fi
[[ -n "$os_name" ]] || os_name=unavailable

kernel=$(uname -srvm 2>/dev/null || printf '%s' unavailable)
time_version=unavailable
if [[ -x /usr/bin/time ]]; then
    time_version=$(/usr/bin/time --version 2>&1 | awk 'NR == 1 {print; exit}' || true)
fi
[[ -n "$time_version" ]] || time_version=unavailable

affinity=unavailable
if command -v taskset >/dev/null 2>&1; then
    affinity=$(taskset -pc "$$" 2>/dev/null || true)
fi
[[ -n "$affinity" ]] || affinity=unavailable

{
    printf 'schema=docx-styles-effects-host-v1\n'
    printf 'utc=%s\n' "$(date -u +%Y-%m-%dT%H:%M:%SZ)"
    printf 'pid=%s\n' "$$"
    printf 'kernel=%s\n' "$kernel"
    printf 'os=%s\n' "$os_name"
    printf 'cpu_model=%s\n' "$cpu_model"
    printf 'logical_cpus=%s\n' "$logical_cpus"
    printf 'memory_total_kib=%s\n' "$memory_total_kib"
    printf 'memory_available_kib=%s\n' "$memory_available_kib"
    printf 'load_1m=%s\n' "$load_1m"
    printf 'time_version=%s\n' "$time_version"
    printf 'affinity=%s\n' "$affinity"
} >"$OUTPUT"
