//! Process-local counters used by the filesystem evidence tranche.
//!
//! The baseline deliberately reads Linux's procfs text interfaces instead of
//! depending on a profiler.  Procfs is best-effort: a platform without these
//! files reports `None` in the child result, and the benchmark remains useful
//! for elapsed time and positional-read accounting.
//!
//! The CPU counters are process CPU-time ticks, not whole-machine utilization.
//! A caller that needs process CPU utilization derives it from
//! `(user_cpu_ticks + system_cpu_ticks) / clock_ticks_per_second / elapsed`.

use std::{fs, io};

use serde::{Deserialize, Serialize};

/// Counters sampled from `/proc/self/io`, `/proc/self/stat`, and
/// `/proc/self/status`.
#[derive(Clone, Copy, Debug, Default, Deserialize, Serialize)]
pub(crate) struct Snapshot {
    /// Bytes returned through read-like system calls.
    pub rchar: u64,
    /// Bytes accepted by write-like system calls.
    pub wchar: u64,
    /// Bytes read from storage (as opposed to page cache).
    pub read_bytes: u64,
    /// Bytes written to storage.
    pub write_bytes: u64,
    /// Bytes cancelled before reaching storage.
    pub cancelled_write_bytes: u64,
    /// Number of read-like system calls.
    pub syscr: u64,
    /// Number of write-like system calls.
    pub syscw: u64,
    /// Minor page faults.
    pub minor_faults: u64,
    /// Major page faults.
    pub major_faults: u64,
    /// User-mode process CPU time in procfs clock ticks.
    pub user_cpu_ticks: u64,
    /// Kernel-mode process CPU time in procfs clock ticks.
    pub system_cpu_ticks: u64,
    /// Procfs clock ticks per second used to interpret the CPU tick fields.
    pub clock_ticks_per_second: u64,
    /// Voluntary context switches performed by this process.
    pub voluntary_context_switches: u64,
    /// Involuntary context switches performed by this process.
    pub nonvoluntary_context_switches: u64,
    /// Resident set size in bytes, when exposed by procfs.
    pub rss_bytes: u64,
    /// High-water resident set size in bytes.
    pub peak_rss_bytes: u64,
}

/// Saturating difference between two process snapshots.
#[derive(Clone, Copy, Debug, Default, Deserialize, Serialize)]
pub(crate) struct Delta {
    pub rchar: u64,
    pub wchar: u64,
    pub read_bytes: u64,
    pub write_bytes: u64,
    pub cancelled_write_bytes: u64,
    pub syscr: u64,
    pub syscw: u64,
    pub minor_faults: u64,
    pub major_faults: u64,
    /// User-mode process CPU time consumed during the operation, in procfs
    /// clock ticks.
    pub user_cpu_ticks: u64,
    /// Kernel-mode process CPU time consumed during the operation, in procfs
    /// clock ticks.
    pub system_cpu_ticks: u64,
    /// The after-sample procfs clock ticks per second used to interpret the
    /// CPU tick deltas.
    pub clock_ticks_per_second: u64,
    /// Voluntary context switches during the operation.
    pub voluntary_context_switches: u64,
    /// Involuntary context switches during the operation.
    pub nonvoluntary_context_switches: u64,
    pub rss_bytes: u64,
    /// The after-sample VmHWM value (not a delta).
    pub peak_rss_bytes: u64,
}

impl Snapshot {
    /// Reads all supported procfs counters for the current process.
    pub(crate) fn read() -> io::Result<Self> {
        let io_text = fs::read_to_string("/proc/self/io")?;
        let stat_text = fs::read_to_string("/proc/self/stat")?;
        let status_text = fs::read_to_string("/proc/self/status")?;
        Ok(Self {
            rchar: key_value(&io_text, "rchar")?,
            wchar: key_value(&io_text, "wchar")?,
            read_bytes: key_value(&io_text, "read_bytes")?,
            write_bytes: key_value(&io_text, "write_bytes")?,
            cancelled_write_bytes: key_value(&io_text, "cancelled_write_bytes")?,
            syscr: key_value(&io_text, "syscr")?,
            syscw: key_value(&io_text, "syscw")?,
            minor_faults: stat_field(&stat_text, 10)?,
            major_faults: stat_field(&stat_text, 12)?,
            user_cpu_ticks: stat_field(&stat_text, 14)?,
            system_cpu_ticks: stat_field(&stat_text, 15)?,
            clock_ticks_per_second: clock_ticks_per_second()?,
            voluntary_context_switches: key_value(&status_text, "voluntary_ctxt_switches")?,
            nonvoluntary_context_switches: key_value(&status_text, "nonvoluntary_ctxt_switches")?,
            rss_bytes: rss_bytes(&status_text)?,
            peak_rss_bytes: vm_hwm_bytes(&status_text)?,
        })
    }

    /// Returns a component-wise saturating difference.
    #[must_use]
    pub(crate) fn delta(self, before: Self) -> Delta {
        Delta {
            rchar: self.rchar.saturating_sub(before.rchar),
            wchar: self.wchar.saturating_sub(before.wchar),
            read_bytes: self.read_bytes.saturating_sub(before.read_bytes),
            write_bytes: self.write_bytes.saturating_sub(before.write_bytes),
            cancelled_write_bytes: self
                .cancelled_write_bytes
                .saturating_sub(before.cancelled_write_bytes),
            syscr: self.syscr.saturating_sub(before.syscr),
            syscw: self.syscw.saturating_sub(before.syscw),
            minor_faults: self.minor_faults.saturating_sub(before.minor_faults),
            major_faults: self.major_faults.saturating_sub(before.major_faults),
            user_cpu_ticks: self.user_cpu_ticks.saturating_sub(before.user_cpu_ticks),
            system_cpu_ticks: self
                .system_cpu_ticks
                .saturating_sub(before.system_cpu_ticks),
            clock_ticks_per_second: self.clock_ticks_per_second,
            voluntary_context_switches: self
                .voluntary_context_switches
                .saturating_sub(before.voluntary_context_switches),
            nonvoluntary_context_switches: self
                .nonvoluntary_context_switches
                .saturating_sub(before.nonvoluntary_context_switches),
            rss_bytes: self.rss_bytes.saturating_sub(before.rss_bytes),
            peak_rss_bytes: self.peak_rss_bytes,
        }
    }
}

#[cfg(target_os = "linux")]
fn clock_ticks_per_second() -> io::Result<u64> {
    let ticks = rustix::param::clock_ticks_per_second();
    if ticks == 0 {
        return Err(io::Error::new(
            io::ErrorKind::InvalidData,
            "clock ticks per second is zero",
        ));
    }
    Ok(ticks)
}

#[cfg(not(target_os = "linux"))]
fn clock_ticks_per_second() -> io::Result<u64> {
    Err(io::Error::new(
        io::ErrorKind::Unsupported,
        "procfs CPU clock ticks are only supported on Linux",
    ))
}

fn key_value(text: &str, key: &str) -> io::Result<u64> {
    let value = text
        .lines()
        .find_map(|line| {
            let (name, value) = line.split_once(':')?;
            (name.trim() == key).then_some(value.trim())
        })
        .ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidData,
                format!("missing /proc key {key}"),
            )
        })?;
    value.parse::<u64>().map_err(|error| {
        io::Error::new(
            io::ErrorKind::InvalidData,
            format!("invalid /proc key {key} value {value:?}: {error}"),
        )
    })
}

/// Parse a field from `/proc/self/stat`. `field` uses the procfs one-based
/// numbering, where field 1 is the process id. The executable name is allowed
/// to contain spaces and parentheses, so parsing starts after the final `)`.
fn stat_field(text: &str, field: usize) -> io::Result<u64> {
    let close = text.rfind(')').ok_or_else(|| {
        io::Error::new(
            io::ErrorKind::InvalidData,
            "malformed /proc/self/stat command",
        )
    })?;
    let index = field.checked_sub(3).ok_or_else(|| {
        io::Error::new(
            io::ErrorKind::InvalidInput,
            "stat field precedes command name",
        )
    })?;
    text.get(close + 1..)
        .and_then(|suffix| suffix.split_whitespace().nth(index))
        .ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidData,
                format!("missing /proc/self/stat field {field}"),
            )
        })?
        .parse::<u64>()
        .map_err(|error| {
            io::Error::new(
                io::ErrorKind::InvalidData,
                format!("invalid /proc/self/stat field {field}: {error}"),
            )
        })
}

fn rss_bytes(text: &str) -> io::Result<u64> {
    kib_bytes(text, "VmRSS")
}

fn vm_hwm_bytes(text: &str) -> io::Result<u64> {
    kib_bytes(text, "VmHWM")
}

fn kib_bytes(text: &str, key: &str) -> io::Result<u64> {
    let kib = text
        .lines()
        .find_map(|line| {
            let (name, value) = line.split_once(':')?;
            if name.trim() != key {
                return None;
            }
            value.split_whitespace().next()
        })
        .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidData, format!("missing /proc {key}")))?
        .parse::<u64>()
        .map_err(|error| {
            io::Error::new(
                io::ErrorKind::InvalidData,
                format!("invalid /proc {key}: {error}"),
            )
        })?;
    kib.checked_mul(1024).ok_or_else(|| {
        io::Error::new(
            io::ErrorKind::InvalidData,
            format!("/proc {key} overflows bytes"),
        )
    })
}

#[cfg(test)]
mod tests {
    use super::{Delta, Snapshot, key_value, rss_bytes, stat_field, vm_hwm_bytes};

    #[test]
    fn parses_stat_after_parenthesized_command() {
        let stat = "7 (name with ) parens) S 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24";
        assert_eq!(stat_field(stat, 10).unwrap(), 7);
        assert_eq!(stat_field(stat, 12).unwrap(), 9);
        assert_eq!(stat_field(stat, 14).unwrap(), 11);
        assert_eq!(stat_field(stat, 15).unwrap(), 12);
        assert_eq!(stat_field(stat, 24).unwrap(), 21);
    }

    #[test]
    fn parses_status_keys_around_parenthesized_values() {
        let status = "Name:\tworker (nested)\n\
            voluntary_ctxt_switches:\t17\n\
            nonvoluntary_ctxt_switches:\t19\n\
            VmRSS:\t23 kB\n";
        assert_eq!(key_value(status, "voluntary_ctxt_switches").unwrap(), 17);
        assert_eq!(key_value(status, "nonvoluntary_ctxt_switches").unwrap(), 19);
    }

    #[test]
    fn rejects_overflowing_proc_counter() {
        let status = "voluntary_ctxt_switches: 18446744073709551616\n";
        assert!(key_value(status, "voluntary_ctxt_switches").is_err());
    }

    #[test]
    fn delta_saturates_counters_and_keeps_after_clock_rate() {
        let before = Snapshot {
            user_cpu_ticks: 50,
            system_cpu_ticks: 40,
            clock_ticks_per_second: 100,
            voluntary_context_switches: 11,
            nonvoluntary_context_switches: 13,
            ..Snapshot::default()
        };
        let after = Snapshot {
            user_cpu_ticks: 75,
            system_cpu_ticks: 30,
            clock_ticks_per_second: 250,
            voluntary_context_switches: 19,
            nonvoluntary_context_switches: 12,
            ..Snapshot::default()
        };
        let delta = after.delta(before);
        assert_eq!(delta.user_cpu_ticks, 25);
        assert_eq!(delta.system_cpu_ticks, 0);
        assert_eq!(delta.clock_ticks_per_second, 250);
        assert_eq!(delta.voluntary_context_switches, 8);
        assert_eq!(delta.nonvoluntary_context_switches, 0);
    }

    #[test]
    fn delta_serializes_cpu_units_explicitly() {
        let delta = Delta {
            user_cpu_ticks: 2,
            system_cpu_ticks: 3,
            clock_ticks_per_second: 100,
            voluntary_context_switches: 4,
            nonvoluntary_context_switches: 5,
            ..Delta::default()
        };
        let json = serde_json::to_value(delta).unwrap();
        assert_eq!(json["user_cpu_ticks"], 2);
        assert_eq!(json["system_cpu_ticks"], 3);
        assert_eq!(json["clock_ticks_per_second"], 100);
        assert_eq!(json["voluntary_context_switches"], 4);
        assert_eq!(json["nonvoluntary_context_switches"], 5);
    }

    #[test]
    fn parses_rss_kib() {
        assert_eq!(rss_bytes("Name: test\nVmRSS: 12 kB\n").unwrap(), 12 * 1024);
        assert_eq!(
            vm_hwm_bytes("Name: test\nVmHWM: 13 kB\n").unwrap(),
            13 * 1024
        );
    }
}
