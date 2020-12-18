#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <sys/stat.h>
#include <unistd.h>

#include "xlstext.h"

#include "../libxls/include/xls.h"

int main(int argc, char *argv[])
{
    set_console();

    Command command;
    {
        Result result = set_command(&command, argv + 1, argc - 1);
        if (result.code)
        {
            const char *app_name = strrchr(argv[0], '\\') ? strrchr(argv[0], '\\') + 1 : strrchr(argv[0], '/') ? strrchr(argv[0], '/') + 1 : argv[0];
            fprintf(stderr, "%s\n", result.desc);
            fprintf(stderr, "Usage: %s <input> [-v] [-csv outpath] [-json outpath]\n", app_name);
            fprintf(stderr, "    <input>          : input xls file or directory\n");
            fprintf(stderr, "    -v               : print log\n");
            fprintf(stderr, "    [-csv output]    : output to csv file or directory\n");
            fprintf(stderr, "    [-json output]   : output to json file or directory\n\n");

            free_result(&result);
            free_command(&command);

            reset_console();
            return -1;
        }
    }

    File *files = NULL;
    int count = 0;
    int is_directory = 0;
    {
        Result result = set_files(&files, &count, &is_directory, command);
        if (result.code)
        {
            char *desc = local_to_utf8(result.desc);
            fprintf(stderr, "%s\n", desc);
            free(desc);

            free_result(&result);
            free_files(&files, count);
            free_command(&command);

            reset_console();
            return -1;
        }
    }

    char *csv_path = NULL, *json_path = NULL;
    for (int i = 0; i < 2; ++i)
    {
        if ((i == 0 && !command.csv) || (i == 1 && !command.json))
            continue;

        char *path = (char *)malloc(strlen(i == 0 ? command.csv : command.json) + 2);
        if (i == 0)
            csv_path = path;
        else
            json_path = path;
        strcpy(path, i == 0 ? command.csv : command.json);
        path[strlen(path) + 1] = 0;

        struct stat st;
        if ((!strcmp(path, ".") || !strcmp(path, "..")) || ((!stat(path, &st) && S_ISDIR(st.st_mode) || (i == 0 && is_directory)) && path[strlen(path) - 1] != '\\' && path[strlen(path) - 1] != '/'))
            path[strlen(path)] = '/';

        char *dot = path;
        while (*dot)
        {
            if (*dot == '/' || *dot == '\\')
            {
                char ch = *dot;
                *dot = 0;
                if (access(path, F_OK))
                {
#ifdef _WIN32
                    if (mkdir(path))
#else
                    if (mkdir(path, 0777))
#endif
                        break;
                }
                *dot = ch;
            }
            ++dot;
        }

        if (i == 1 && csv_path && csv_path[strlen(csv_path) - 1] != '\\' && csv_path[strlen(csv_path) - 1] != '/' && !stat(csv_path, &st) && S_ISDIR(st.st_mode))
            csv_path[strlen(csv_path)] = '/';
    }

    FILE *jsons = NULL;
    int ok_cnt = 0;
    for (int i = 0; i < count; ++i)
    {
        const char *path = files[i].path, *name = files[i].name;

        FILE *csv = NULL, *json = NULL;
        char result[100] = "";
        do
        {
            xls_error_t error = LIBXLS_OK;
            xlsWorkBook *workbook = xls_open_file(path, "UTF-8", &error);
            if (error != LIBXLS_OK)
            {
                strcpy(result, xls_getError(error));
                break;
            }
            xlsWorkSheet *worksheet = xls_getWorkSheet(workbook, 0);
            if (!worksheet)
            {
                strcpy(result, "no a sheet");
                break;
            }
            error = xls_parseWorkSheet(worksheet);
            if (error != LIBXLS_OK)
            {
                strcpy(result, xls_getError(error));
                break;
            }

            for (int i = 0; i < 2; ++i)
            {
                if ((i == 0 && !csv_path) || (i == 1 && !json_path))
                    continue;

                char *url = i == 0 ? csv_path : json_path;
                int is_out_directory = url[strlen(url) - 1] == '\\' || url[strlen(url) - 1] == '/';
                if (is_out_directory)
                {
                    char *url_tmp = (char *)malloc(strlen(url) + strlen(name) + 5);
                    strcpy(url_tmp, url);
                    strcat(url_tmp, name);
                    char *dot = strrchr(url_tmp, '.');
                    const char *ext = i == 0 ? ".csv" : ".json";
                    if (dot && (dot[1] == 'x' || dot[1] == 'X') && (dot[2] == 'l' || dot[2] == 'L') && (dot[3] == 's' || dot[3] == 'S') && dot[4] == 0)
                        strcpy(dot, ext);
                    else
                        strcat(url_tmp, ext);
                    url = url_tmp;
                }
                else if (i == 1 && !jsons)
                {
                    jsons = fopen(url, "w");
                }

                if (i == 0)
                    csv = fopen(url, "w");
                else
                    json = (is_out_directory ? fopen(url, "w") : jsons);
                if (is_out_directory)
                    free(url);
            }
            if ((csv_path && !csv) || (json_path && !json))
            {
                strcpy(result, !csv && !json ? "cannot create csv and json file" : !csv ? "cannot create csv file" : "cannot create json file");
                break;
            }

            if (json)
            {
                if (json == jsons)
                {
                    char *shortname = local_to_utf8(name);
                    char *dot = strrchr(shortname, '.');
                    if (dot && (dot[1] == 'x' || dot[1] == 'X') && (dot[2] == 'l' || dot[2] == 'L') && (dot[3] == 's' || dot[3] == 'S') && dot[4] == 0)
                        *dot = 0;

                    char *name = to_json(shortname);
                    fprintf(json, "%s    %s: [\n", i > 0 ? ",\n" : "{\n", name);
                    free(name);
                    free(shortname);
                }
                else
                    fputs("[\n", json);
            }

            WORD *spans = NULL;
            int spans_count = 0;
            for (WORD row = 0; row <= worksheet->rows.lastrow; ++row)
            {
                if (row > 0)
                {
                    if (json)
                        fputs(",\n", json);
                    if (csv)
                        fputs("\n", csv);
                }
                if (json)
                    fputs(json == jsons ? "        [" : "    [", json);

                for (WORD col = 0; col <= worksheet->rows.lastcol; ++col)
                {
                    xlsCell *cell = xls_cell(worksheet, row, col);
                    if (cell->rowspan > 0 && cell->colspan > 0)
                    {
                        spans = (WORD *)realloc(spans, sizeof(WORD) * (spans_count + 1) * 4);
                        WORD *span = spans + spans_count * 4;
                        span[0] = row;
                        span[1] = col;
                        span[2] = cell->rowspan;
                        span[3] = cell->colspan;
                        ++spans_count;
                    }
                    else
                    {
                        for (int i = 0; i < spans_count; ++i)
                        {
                            WORD *span = spans + i * 4;
                            if (span[0] <= row && row < span[0] + span[2] && span[1] <= col && col < span[1] + span[3])
                            {
                                cell = xls_cell(worksheet, span[0], span[1]);
                                break;
                            }
                        }
                    }

                    char *value = strtrim(strdup(cell->str ? cell->str : ""));
                    int is_str = 1;
                    if (cell->id == XLS_RECORD_RK || cell->id == XLS_RECORD_MULRK || cell->id == XLS_RECORD_NUMBER)
                    {
                        sprintf(value, "%.15g", cell->d);
                        is_str = 0;
                    }
                    else if (cell->id == XLS_RECORD_FORMULA || cell->id == XLS_RECORD_FORMULA_ALT)
                    {
                        if (cell->l == 0)
                        {
                            sprintf(value, "%.15g", cell->d);
                            is_str = 0;
                        }
                        else
                        {
                            if (!strcmp(cell->str, "bool"))
                            {
                                strcpy(value, cell->d ? "true" : "false");
                                is_str = 0;
                            }
                            else if (!strcmp(cell->str, "error"))
                            {
                                *value = 0;
                                is_str = 0;
                            }
                        }
                    }
                    if (csv)
                    {
                        char *csv_value = to_csv(value);
                        fprintf(csv, col > 0 ? ",%s" : "%s", csv_value);
                        free(csv_value);
                    }
                    if (json)
                    {
                        char *json_value = is_str ? to_json(value) : strdup(value[0] ? value : "\"\"");
                        fprintf(json, col > 0 ? ",%s" : "%s", json_value);
                        free(json_value);
                    }
                    free(value);
                }
                if (json)
                    fputc(']', json);
            }
            free(spans);

            if (json)
                fprintf(json, "\n%s]", json == jsons ? "    " : "");

            xls_close_WS(worksheet);
            xls_close_WB(workbook);
        } while (0);

        if (csv)
            fclose(csv);
        if (json && json != jsons)
            fclose(json);

        if (command.verbose)
        {
            int is_ok = *result == 0;
            ok_cnt += is_ok ? 1 : 0;
            char *name_utf8 = local_to_utf8(name);
            fprintf(is_ok ? stdout : stderr, "%s (%d/%d): %s%s%s\n", is_ok ? "√" : "×", i + 1, count, name_utf8, is_ok ? "" : ": ", result);
            free(name_utf8);
        }
    }
    if (command.verbose)
        fprintf(stdout, "Result: ok: %d, error: %d.\n", ok_cnt, count - ok_cnt);
    if (jsons)
    {
        fputs("\n}", jsons);
        fclose(jsons);
    }

    free(csv_path);
    free(json_path);
    free_files(&files, count);
    free_command(&command);

    reset_console();
    return 0;
}
