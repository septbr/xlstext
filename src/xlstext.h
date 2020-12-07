void set_console();
void reset_console();

char *strtrim(char *str);
char *local_to_utf8(const char *in);

typedef struct
{
    int code;
    char *desc;
} Result;
Result create_result(int code, const char *desc);
void free_result(Result *result);

typedef struct
{
    char *input; // required
    int verbose; // optional, default 0
    char *csv;   // optional
    char *json;  // optional
} Command;
/**
 * <input> [-v] [-csv outpath] [-json outpath]
 */
Result set_command(Command *command, char *argv[], int argc);
void free_command(Command *command);

typedef struct
{
    char *path;
    char *name;
} File;

/**
 * read xls files
 * @returns count of files
 */
Result set_files(File *files[], int *count, int *is_directory, Command command);
void free_files(File *files[], int count);

/**
 * * MS-DOS-style lines that end with (CR/LF) characters (optional for the last line)
 * * An optional header record (there is no sure way to detect whether it is present, so care is required when importing).
 * * Each record "should" contain the same number of comma-separated fields.
 * * Any field may be quoted (with double quotes).
 * * Fields containing a line-break, double-quote, and/or commas should be quoted. (If they are not, the file will likely be impossible to process correctly).
 * * A (double) quote character in a field must be represented by two (double) quote characters.
 */
char *to_csv(const char *str);

/**
 * \b \f \r \n \t \\ \"
 */
char *to_json(const char *str);
