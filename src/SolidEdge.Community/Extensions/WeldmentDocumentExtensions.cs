using System;
using SolidEdgeFramework;
using SolidEdgePart;

namespace SolidEdgeCommunity.Extensions;

/// <summary>
///     SolidEdgePart.WeldmentDocument extension methods.
/// </summary>
public static class WeldmentDocumentExtensions
{
    /// <summary>
    ///     Returns the version of Solid Edge that was used to create the referenced document.
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    public static Version GetCreatedVersion(this WeldmentDocument document)
    {
        return new Version(document.CreatedVersion);
    }

    /// <summary>
    ///     Returns the version of Solid Edge that was used the last time the referenced document was saved.
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    public static Version GetLastSavedVersion(this WeldmentDocument document)
    {
        return new Version(document.LastSavedVersion);
    }

    /// <summary>
    ///     Returns the properties for the referenced document.
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    public static PropertySets GetProperties(this WeldmentDocument document)
    {
        return document.Properties as PropertySets;
    }

    /// <summary>
    ///     Returns the summary information property set for the referenced document.
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    public static SummaryInfo GetSummaryInfo(this WeldmentDocument document)
    {
        return document.SummaryInfo as SummaryInfo;
    }

    /// <summary>
    ///     Returns a collection of variables for the referenced document.
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    public static Variables GetVariables(this WeldmentDocument document)
    {
        return document.Variables as Variables;
    }
}