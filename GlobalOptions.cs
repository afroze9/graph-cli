using System.CommandLine;

namespace GraphCli;

public static class GlobalOptions
{
    public static readonly Option<string> Format = new("--format")
    {
        Description = "Output format: json or table",
        DefaultValueFactory = _ => "json",
        Recursive = true
    };

    public static readonly Option<string?> TimeZone = new("--timezone")
    {
        Description = "Timezone for datetime input/output (IANA e.g. 'Asia/Karachi' or Windows e.g. 'Pakistan Standard Time'). Defaults to local system timezone.",
        Recursive = true
    };
}
