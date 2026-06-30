#ifndef __SLOG_H__

#ifdef _WIN32
#  ifdef SLOG_EXPORTS
#    define SLOG_API __declspec(dllexport)
#  else
#    define SLOG_API __declspec(dllimport)
#  endif
#else
#  define SLOG_API
#endif

#ifdef __cplusplus
extern "C" {
#endif
SLOG_API int slog_init(char *log_cfg, int len);
SLOG_API int slog_format(char *buf, int buf_size);
SLOG_API int slog_parse(char *data, int len, char **result, int *updated);
SLOG_API int slog_free(void);
#ifdef __cplusplus
}
#endif

#endif