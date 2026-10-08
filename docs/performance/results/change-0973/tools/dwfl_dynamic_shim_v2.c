#define _GNU_SOURCE

#include <dlfcn.h>
#include <elfutils/libdwfl.h>
#include <errno.h>
#include <fcntl.h>
#include <inttypes.h>
#include <stdbool.h>
#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <unistd.h>

typedef Dwfl_Module *(*dwfl_report_elf_fn)(Dwfl *, const char *, const char *,
                                            int, GElf_Addr, bool);
typedef const char *(*dwfl_module_info_fn)(Dwfl_Module *, void ***, Dwarf_Addr *,
                                            Dwarf_Addr *, Dwarf_Addr *,
                                            Dwarf_Addr *, const char **,
                                            const char **);
typedef int (*dwfl_errno_fn)(void);

static dwfl_report_elf_fn real_report;
static dwfl_module_info_fn real_info;
static dwfl_errno_fn real_errno;
static int log_fd = -1;
static __thread int resolving;

static void *resolve_symbol(const char *name, const char *version) {
    void *value = dlvsym(RTLD_NEXT, name, version);
    if (value == NULL)
        value = dlsym(RTLD_NEXT, name);
    return value;
}

static dwfl_report_elf_fn resolve_report(void) {
    if (real_report != NULL)
        return real_report;
    if (resolving)
        return NULL;
    resolving = 1;
    real_report = (dwfl_report_elf_fn)
        resolve_symbol("dwfl_report_elf", "ELFUTILS_0.156");
    resolving = 0;
    return real_report;
}

static dwfl_module_info_fn resolve_info(void) {
    if (real_info != NULL)
        return real_info;
    if (resolving)
        return NULL;
    resolving = 1;
    real_info = (dwfl_module_info_fn)
        resolve_symbol("dwfl_module_info", "ELFUTILS_0.122");
    resolving = 0;
    return real_info;
}

static dwfl_errno_fn resolve_errno(void) {
    if (real_errno != NULL)
        return real_errno;
    resolving = 1;
    real_errno = (dwfl_errno_fn)
        resolve_symbol("dwfl_errno", "ELFUTILS_0.122");
    resolving = 0;
    return real_errno;
}

static const char *mode_name(void) {
    const char *mode = getenv("LITCHI_0973_SHIM_MODE");
    if (mode != NULL && (strcmp(mode, "noop") == 0 || strcmp(mode, "corrected") == 0))
        return mode;
    return "invalid";
}

static const char *configuration(const char *key) {
    const char *value = getenv(key);
    return value != NULL ? value : "";
}

static bool parse_address(const char *key, GElf_Addr *out) {
    const char *text = configuration(key);
    if (*text == '\0')
        return false;
    char *end = NULL;
    errno = 0;
    unsigned long long value = strtoull(text, &end, 0);
    if (errno != 0 || end == text || *end != '\0')
        return false;
    *out = (GElf_Addr)value;
    return true;
}

static bool config_valid(GElf_Addr *corrected_base, GElf_Addr *perf_expected_start) {
    return configuration("LITCHI_0973_TARGET_PATH")[0] != '\0' &&
           configuration("LITCHI_0973_TARGET_NAME")[0] != '\0' &&
           parse_address("LITCHI_0973_CORRECTED_BASE", corrected_base) &&
           parse_address("LITCHI_0973_PERF_EXPECTED_START", perf_expected_start);
}

static int output_fd(void) {
    if (log_fd >= 0)
        return log_fd;
    const char *path = getenv("LITCHI_0973_SHIM_LOG");
    if (path == NULL || *path == '\0')
        return -1;
    int fd = open(path, O_WRONLY | O_CREAT | O_APPEND | O_CLOEXEC, 0644);
    if (fd >= 0)
        log_fd = fd;
    return fd;
}

static int caller_is_perf_from_pc(void *pc, const char **caller_file,
                                  const char **caller_symbol) {
    Dl_info info;
    memset(&info, 0, sizeof(info));
    int found = dladdr(pc, &info);
    if (caller_file != NULL)
        *caller_file = found && info.dli_fname != NULL ? info.dli_fname : "";
    if (caller_symbol != NULL)
        *caller_symbol = found && info.dli_sname != NULL ? info.dli_sname : "";
    return found && info.dli_fname != NULL &&
           strcmp(info.dli_fname, "/usr/bin/perf") == 0;
}

static void log_report(const char *mode, const char *name, const char *file_name,
                       int fd, GElf_Addr original, GElf_Addr forwarded,
                       bool add_p_vaddr, int target_match,
                       int replacement_applied, Dwfl_Module *result,
                       int saved_errno, int dwfl_error, int config_ok) {
    int out = output_fd();
    if (out < 0)
        return;
    dprintf(out,
            "{\"schema\":\"litchi.performance.0973.dwfl_report.v1\","
            "\"mode\":\"%s\",\"name\":\"%s\","
            "\"file_name\":\"%s\",\"fd\":%d,"
            "\"target_match\":%s,\"config_valid\":%s,"
            "\"original_base\":\"0x%llx\","
            "\"forwarded_base\":\"0x%llx\","
            "\"add_p_vaddr\":%s,\"replacement_applied\":%s,"
            "\"result_module\":\"%p\",\"errno_after\":%d,"
            "\"dwfl_errno_after\":%d}\n",
            mode, name != NULL ? name : "", file_name != NULL ? file_name : "",
            fd, target_match ? "true" : "false", config_ok ? "true" : "false",
            (unsigned long long)original, (unsigned long long)forwarded,
            add_p_vaddr ? "true" : "false",
            replacement_applied ? "true" : "false", (void *)result,
            saved_errno, dwfl_error);
}

static void log_info(const char *mode, const char *module_name,
                     const char *mainfile, const char *debugfile,
                     Dwarf_Addr before_start, Dwarf_Addr after_start,
                     Dwarf_Addr end, Dwarf_Addr dwbias, Dwarf_Addr symbias,
                     int target_match, int caller_perf,
                     const char *caller_file, const char *caller_symbol,
                     void *caller_pc, int adjusted, int config_ok) {
    int out = output_fd();
    if (out < 0)
        return;
    dprintf(out,
            "{\"schema\":\"litchi.performance.0973.dwfl_info.v1\","
            "\"mode\":\"%s\",\"module_name\":\"%s\","
            "\"mainfile\":\"%s\",\"debugfile\":\"%s\","
            "\"before_start\":\"0x%llx\",\"after_start\":\"0x%llx\","
            "\"end\":\"0x%llx\",\"dwbias\":\"0x%llx\","
            "\"symbias\":\"0x%llx\",\"target_match\":%s,"
            "\"caller_is_perf\":%s,\"caller_file\":\"%s\","
            "\"caller_symbol\":\"%s\",\"caller_pc\":\"%p\","
            "\"start_adjusted\":%s,\"config_valid\":%s}\n",
            mode, module_name != NULL ? module_name : "",
            mainfile != NULL ? mainfile : "", debugfile != NULL ? debugfile : "",
            (unsigned long long)before_start, (unsigned long long)after_start,
            (unsigned long long)end, (unsigned long long)dwbias,
            (unsigned long long)symbias, target_match ? "true" : "false",
            caller_perf ? "true" : "false", caller_file != NULL ? caller_file : "",
            caller_symbol != NULL ? caller_symbol : "", caller_pc,
            adjusted ? "true" : "false", config_ok ? "true" : "false");
}

Dwfl_Module *dwfl_report_elf(Dwfl *dwfl, const char *name, const char *file_name,
                             int fd, GElf_Addr base, bool add_p_vaddr) {
    dwfl_report_elf_fn next = resolve_report();
    if (next == NULL)
        return NULL;
    GElf_Addr corrected_base = 0;
    GElf_Addr perf_expected_start = 0;
    int config_ok = config_valid(&corrected_base, &perf_expected_start);
    (void)perf_expected_start;
    const char *target_path = configuration("LITCHI_0973_TARGET_PATH");
    int target_match = *target_path != '\0' && file_name != NULL &&
                       strcmp(file_name, target_path) == 0;
    const char *mode = mode_name();
    int replace = target_match && config_ok && strcmp(mode, "corrected") == 0;
    GElf_Addr forwarded = replace ? corrected_base : base;
    Dwfl_Module *result = next(dwfl, name, file_name, fd, forwarded, add_p_vaddr);
    int saved_errno = errno;
    int error = -1;
    dwfl_errno_fn get_errno = resolve_errno();
    if (get_errno != NULL)
        error = get_errno();
    log_report(mode, name, file_name, fd, base, forwarded, add_p_vaddr,
               target_match, replace, result, saved_errno, error, config_ok);
    errno = saved_errno;
    return result;
}

const char *dwfl_module_info(Dwfl_Module *mod, void ***userdata,
                             Dwarf_Addr *start, Dwarf_Addr *end,
                             Dwarf_Addr *dwbias, Dwarf_Addr *symbias,
                             const char **mainfile, const char **debugfile) {
    dwfl_module_info_fn next = resolve_info();
    if (next == NULL)
        return NULL;
    const char *name = next(mod, userdata, start, end, dwbias, symbias,
                            mainfile, debugfile);
    int saved_errno = errno;
    Dwarf_Addr before = start != NULL ? *start : 0;
    Dwarf_Addr after = before;
    const char *caller_file = "";
    const char *caller_symbol = "";
    void *caller_pc = __builtin_return_address(0);
    int direct_perf = caller_is_perf_from_pc(caller_pc, &caller_file,
                                             &caller_symbol);
    const char *target_name = configuration("LITCHI_0973_TARGET_NAME");
    int target_match = *target_name != '\0' && name != NULL &&
                       strcmp(name, target_name) == 0;
    GElf_Addr corrected_base = 0;
    GElf_Addr perf_expected_start = 0;
    int config_ok = config_valid(&corrected_base, &perf_expected_start);
    int adjusted = 0;
    if (direct_perf && target_match && config_ok && start != NULL &&
        *start == corrected_base && strcmp(mode_name(), "corrected") == 0) {
        *start = perf_expected_start;
        after = *start;
        adjusted = 1;
    }
    if (start != NULL)
        after = *start;
    log_info(mode_name(), name, mainfile != NULL ? *mainfile : NULL,
             debugfile != NULL ? *debugfile : NULL, before, after,
             end != NULL ? *end : 0, dwbias != NULL ? *dwbias : 0,
             symbias != NULL ? *symbias : 0, target_match, direct_perf,
             caller_file, caller_symbol, caller_pc, adjusted, config_ok);
    errno = saved_errno;
    return name;
}
