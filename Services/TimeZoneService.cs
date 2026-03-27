using Microsoft.Graph.Models;

namespace GraphCli.Services;

public static class TimeZoneService
{
    /// <summary>
    /// Resolves a timezone string to a Windows timezone ID suitable for the Graph API.
    /// Accepts Windows IDs (e.g., "Pakistan Standard Time") or IANA IDs (e.g., "Asia/Karachi").
    /// Returns the local timezone ID if input is null/empty.
    /// </summary>
    public static string ResolveTimeZoneId(string? timezone)
    {
        if (string.IsNullOrWhiteSpace(timezone))
            return TimeZoneInfo.Local.Id;

        // Try as a Windows timezone ID first
        try
        {
            var tz = TimeZoneInfo.FindSystemTimeZoneById(timezone);
            return tz.Id;
        }
        catch (TimeZoneNotFoundException)
        {
            throw new ArgumentException(
                $"Unknown timezone: '{timezone}'. Use a valid IANA ID (e.g., 'Asia/Karachi') " +
                $"or Windows ID (e.g., 'Pakistan Standard Time').");
        }
    }

    /// <summary>
    /// Converts a DateTimeTimeZone from the Graph API to the target timezone,
    /// returning the datetime string in the target timezone.
    /// </summary>
    public static string? ConvertToTimeZone(string? dateTimeString, string? sourceTimeZone, string targetTimeZoneId)
    {
        if (string.IsNullOrEmpty(dateTimeString))
            return dateTimeString;

        if (!DateTime.TryParse(dateTimeString, out var dt))
            return dateTimeString;

        try
        {
            var sourceId = string.IsNullOrEmpty(sourceTimeZone) ? "UTC" : sourceTimeZone;
            var sourceTz = TimeZoneInfo.FindSystemTimeZoneById(sourceId);
            var targetTz = TimeZoneInfo.FindSystemTimeZoneById(targetTimeZoneId);

            var sourceTime = DateTime.SpecifyKind(dt, DateTimeKind.Unspecified);
            var converted = TimeZoneInfo.ConvertTime(sourceTime, sourceTz, targetTz);
            return converted.ToString("yyyy-MM-ddTHH:mm:ss.0000000");
        }
        catch
        {
            // If conversion fails, return the original
            return dateTimeString;
        }
    }

    /// <summary>
    /// Converts a DateTimeOffset (e.g., from mail ReceivedDateTime) to the target timezone.
    /// </summary>
    public static DateTimeOffset? ConvertToTimeZone(DateTimeOffset? dateTimeOffset, string targetTimeZoneId)
    {
        if (dateTimeOffset == null)
            return null;

        try
        {
            var targetTz = TimeZoneInfo.FindSystemTimeZoneById(targetTimeZoneId);
            return TimeZoneInfo.ConvertTime(dateTimeOffset.Value, targetTz);
        }
        catch
        {
            return dateTimeOffset;
        }
    }
}
