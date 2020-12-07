#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <sys/stat.h>

#include <dirent.h>
#include <iconv.h>

#include "xlstext.h"

#ifdef WIN32
#include <Windows.h>
int init = 0;
UINT console_cp = 0;
#endif

void set_console()
{
#ifdef _WIN32
    if (!init)
    {
        init = 1;
        console_cp = GetConsoleOutputCP();
        SetConsoleOutputCP(CP_UTF8);
    }
#endif
}
void reset_console()
{
#ifdef _WIN32
    if (init)
    {
        init = 0;
        SetConsoleOutputCP(console_cp);
    }
#endif
}

char *strtrim(char *str)
{
    char *from = str, *end = str + strlen(str) - 1;
    while (*from && isspace(*from))
        ++from;
    while (*end && isspace(*end))
        --end;
    end[1] = 0;
    strcpy(str, from);
    return str;
}

char *local_to_utf8(const char *in)
{
#ifdef _WIN32
    size_t len_in = strlen(in);
    size_t len_out = len_in * 5;

    char *out = malloc(len_out);
    char *pout = out;

    iconv_t cv = iconv_open("utf-8", "");
    iconv(cv, (char **)&in, &len_in, &pout, &len_out);
    iconv_close(cv);
    *pout = 0;

    return out;
#else
    return strdup(in);
#endif
}

Result create_result(int code, const char *desc)
{
    Result result = {code, desc ? strdup(desc) : NULL};
    return result;
}
void free_result(Result *result) { free(result->desc); }

/**
 * <input> [-v] [-csv outpath] [-json outpath]
 */
Result set_command(Command *command, char *argv[], int argc)
{
    memset(command, 0, sizeof(Command));
    int verbose = -1, json_merge = -1;

    for (int i = 0; i < argc; i++)
    {
        char *arg = strtrim(strdup(argv[i]));
        if (!strcmp(arg, "-v"))
        {
            if (verbose != -1)
            {
                free(arg);
                return create_result(-1, "Repeat parameter.");
            }
            command->verbose = verbose = 1;
        }
        else if (!strcmp(arg, "-csv"))
        {
            char *next_arg = ++i < argc ? strtrim(strdup(argv[i])) : NULL;
            if (command->csv || !next_arg || !strlen(next_arg))
            {
                free(next_arg);
                free(arg);
                return create_result(-1, command->csv ? "Repeat parameter." : "Invalid value of \"-csv\".");
            }
            command->csv = next_arg;
        }
        else if (!strcmp(arg, "-json"))
        {
            char *next_arg = ++i < argc ? strtrim(strdup(argv[i])) : NULL;
            if (command->json || !next_arg || !strlen(next_arg))
            {
                free(next_arg);
                free(arg);
                return create_result(-1, command->json ? "Repeat parameter." : "Invalid value of \"-json\".");
            }
            command->json = next_arg;
        }
        else if (arg[0] == '-' || command->input)
        {
            const char *format = "Invalid value of \"%s\".";
            char *desc = (char *)malloc(strlen(format) + strlen(arg) + 1);
            sprintf(desc, format, arg);
            Result result = create_result(-1, desc);
            free(desc);
            free(arg);
            return result;
        }
        else
        {
            command->input = strdup(arg);
        }
        free(arg);
    }

    if (!command->input || (!command->csv && !command->json))
        return create_result(-1, "Too few parameters.");

    return create_result(0, NULL);
}
void free_command(Command *command)
{
    free(command->input);
    command->input = NULL;
    free(command->csv);
    command->csv = NULL;
    free(command->json);
    command->json = NULL;
}

/**
 * read xls files
 * @returns count of files
 */
Result set_files(File *files[], int *count, int *is_directory, Command command)
{
    *files = NULL;
    *count = 0;
    *is_directory = 0;

    if (!command.input)
        return create_result(-1, "no input.");

    struct stat st;
    if (stat(command.input, &st))
    {
        const char *format = "not found: \"%s\".";
        char *desc = (char *)malloc(strlen(format) + strlen(command.input) + 1);
        sprintf(desc, format, command.input);
        Result result = create_result(-1, desc);
        free(desc);
        return result;
    }
    if (S_ISDIR(st.st_mode))
    {
        *is_directory = 1;
        char *dir_path = strdup(command.input);
        if (dir_path[strlen(dir_path) - 1] == '/' || dir_path[strlen(dir_path) - 1] == '\\')
            dir_path[strlen(dir_path) - 1] = 0;

        DIR *dir = opendir(dir_path);
        if (!dir)
        {
            const char *format = "can not read: \"%s\".";
            char *desc = (char *)malloc(strlen(format) + strlen(dir_path) + 1);
            sprintf(desc, format, dir_path);
            Result result = create_result(-1, desc);
            free(desc);
            return result;
        }
        struct dirent *entry;
        while ((entry = readdir(dir)) != NULL)
        {
            char *path = (char *)malloc(strlen(dir_path) + 1 + strlen(entry->d_name) + 1);
            char *name = path + strlen(dir_path) + 1;
            sprintf(path, "%s/%s", dir_path, entry->d_name);

            if (!stat(path, &st) && S_ISREG(st.st_mode))
            {
                const char *ext = strrchr(path, '.');
                if (!stricmp(ext, ".xls"))
                {
                    *files = (File *)realloc(*files, sizeof(File) * (*count + 1));
                    (*files)[*count].path = path;
                    (*files)[*count].name = name;
                    ++*count;
                }
            }
        }
        closedir(dir);
        free(dir_path);

        if (*count == 0)
        {
            const char *format = "not xls files in \"%s\".";
            char *desc = (char *)malloc(strlen(format) + strlen(command.input) + 1);
            sprintf(desc, format, command.input);
            Result result = create_result(-1, desc);
            free(desc);
            return result;
        }
    }
    else //if (S_ISREG(st.st_mode))
    {
        char *path = strdup(command.input);
        while (path[strlen(path) - 1] == '/' || path[strlen(path) - 1] == '\\')
            path[strlen(path) - 1] = 0;
        char *name = strrchr(path, '\\') ? strrchr(path, '\\') + 1 : strrchr(path, '/') ? strrchr(path, '/') + 1 : path;

        *files = (File *)realloc(*files, sizeof(File) * (*count + 1));
        (*files)[*count].path = path;
        (*files)[*count].name = name;
        ++*count;
    }
    return create_result(0, NULL);
}
void free_files(File *files[], int count)
{
    for (int i = 0; i < count; ++i)
        free(files[i]->path);
    free(*files);
    *files = NULL;
}

/**
 * * MS-DOS-style lines that end with (CR/LF) characters (optional for the last line)
 * * An optional header record (there is no sure way to detect whether it is present, so care is required when importing).
 * * Each record "should" contain the same number of comma-separated fields.
 * * Any field may be quoted (with double quotes).
 * * Fields containing a line-break, double-quote, and/or commas should be quoted. (If they are not, the file will likely be impossible to process correctly).
 * * A (double) quote character in a field must be represented by two (double) quote characters.
 */
char *to_csv(const char *str)
{
    const char *ch = str;
    size_t count = strlen(str) + 1, quoted = 0;
    while (*ch)
    {
        if (*ch == '\"' || *ch == '\r' || *ch == '\n' || *ch == ',')
            quoted = 2;
        if (*ch == '\"')
            ++count;
        ++ch;
    }

    char *str_csv = (char *)malloc(count + quoted);
    char *ch_csv = str_csv;
    if (quoted)
        *ch_csv++ = '"';
    ch = str;
    while (*ch)
    {
        if (*ch == '\"')
            *ch_csv++ = '"';
        *ch_csv++ = *ch;
        ++ch;
    }
    if (quoted)
        *ch_csv++ = '"';
    *ch_csv = 0;
    return str_csv;
}

/**
 * \b \f \r \n \t \\ \"
 */
char *to_json(const char *str)
{
    const char *ch = str;
    size_t count = strlen(str) + 1;
    while (*ch)
    {
        if (*ch == '\b' || *ch == '\f' || *ch == '\r' || *ch == '\n' || *ch == '\t' || *ch == '\\' || *ch == '\"')
            ++count;
        else if (iscntrl(*ch))
            --count;
        ++ch;
    }

    char *str_json = (char *)malloc(count + 2);
    char *ch_json = str_json;
    *ch_json++ = '"';
    ch = str;
    while (*ch)
    {
        if (*ch == '\b' || *ch == '\f' || *ch == '\r' || *ch == '\n' || *ch == '\t' || *ch == '\\' || *ch == '\"')
        {
            *ch_json++ = '\\';
            *ch_json++ = *ch == '\b' ? 'b' : *ch == '\f' ? 'f' : *ch == '\r' ? 'r' : *ch == '\n' ? 'n' : *ch == '\t' ? 't' : *ch;
        }
        else if (!iscntrl(*ch))
            *ch_json++ = *ch;
        ++ch;
    }
    *ch_json++ = '"';
    *ch_json = 0;
    return str_json;
}
