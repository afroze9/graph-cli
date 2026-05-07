namespace GraphCli.Services;

public static class UserService
{
    public static async Task<object> GetMeAsync()
    {
        var client = await GraphClientProvider.CreateAsync();
        var me = await client.Me.GetAsync(r =>
        {
            r.QueryParameters.Select = ["id", "displayName", "mail", "userPrincipalName", "jobTitle", "department", "officeLocation", "mobilePhone", "businessPhones"];
        });
        return new
        {
            me!.Id, me.DisplayName, me.Mail, me.UserPrincipalName,
            me.JobTitle, me.Department, me.OfficeLocation,
            me.MobilePhone, me.BusinessPhones
        };
    }

    public static async Task<object> GetUserAsync(string userId)
    {
        var client = await GraphClientProvider.CreateAsync();
        var user = await client.Users[userId].GetAsync(r =>
        {
            r.QueryParameters.Select = ["id", "displayName", "mail", "userPrincipalName", "jobTitle", "department"];
        });
        return new
        {
            user!.Id, user.DisplayName, user.Mail,
            user.UserPrincipalName, user.JobTitle, user.Department
        };
    }

    public static async Task<object> SearchAsync(string query)
    {
        var client = await GraphClientProvider.CreateAsync();
        var escaped = query.Replace("'", "''");
        var users = await client.Users.GetAsync(r =>
        {
            r.QueryParameters.Filter = $"startsWith(displayName,'{escaped}') or startsWith(mail,'{escaped}')";
            r.QueryParameters.Select = ["id", "displayName", "mail", "userPrincipalName", "jobTitle"];
            r.QueryParameters.Top = 25;
        });
        return users?.Value?.Select(u => new
        {
            u.Id, u.DisplayName, u.Mail, u.UserPrincipalName, u.JobTitle
        }).ToList() ?? [];
    }

    public static async Task<object> GetManagerAsync()
    {
        var client = await GraphClientProvider.CreateAsync();
        var manager = await client.Me.Manager.GetAsync();
        if (manager?.Id == null)
            return new { error = "no_manager", message = "No manager set for this account" };

        return await GetUserAsync(manager.Id);
    }

    public static async Task<object> GetPhotoAsync(string userId, string outPath)
    {
        var client = await GraphClientProvider.CreateAsync();
        var stream = await client.Users[userId].Photo.Content.GetAsync();
        if (stream == null)
        {
            return new { status = "error", message = "no photo" };
        }
        await using var fs = File.Create(outPath);
        await stream.CopyToAsync(fs);
        return new { status = "downloaded", file = outPath, size = new FileInfo(outPath).Length };
    }

    public static async Task<object> GetReportsAsync()
    {
        var client = await GraphClientProvider.CreateAsync();
        var reports = await client.Me.DirectReports.GetAsync();
        if (reports?.Value == null || reports.Value.Count == 0)
            return Array.Empty<object>();

        var results = new List<object>();
        foreach (var report in reports.Value)
        {
            if (report.Id != null)
                results.Add(await GetUserAsync(report.Id));
        }
        return results;
    }
}
