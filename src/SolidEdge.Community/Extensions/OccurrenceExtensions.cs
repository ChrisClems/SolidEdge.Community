using System;
using System.Linq;
using SolidEdgeAssembly;
using SolidEdgeFramework;

namespace SolidEdgeCommunity.Extensions;

/// <summary>
///     SolidEdgePart.WeldmentDocument extension methods.
/// </summary>
public static class OccurrenceExtensions
{
    /// <summary>
    ///     Returns the version of Solid Edge that was used to create the referenced document.
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    public static SolidEdgeDocument GetOccurrenceDocument(this Occurrence occurrence)
    {
        return occurrence as SolidEdgeDocument;
    }

    /// <summary>
    ///     Returns a strongly typed occurrence document.
    /// </summary>
    /// <typeparam name="T">The type to return.</typeparam>
    public static T GetOccurrenceDocument<T>(this Occurrence occurrence) where T : class
    {
        return occurrence.GetOccurrenceDocument<T>(true);
    }

    /// <summary>
    ///     Returns a strongly typed occurrence document.
    /// </summary>
    /// <typeparam name="T">The type to return.</typeparam>
    public static T GetOccurrenceDocument<T>(this Occurrence occurrence, bool throwOnError) where T : class
    {
        try
        {
            return (T)occurrence.OccurrenceDocument;
        }
        catch
        {
            if (throwOnError) throw;
        }

        return null;
    }

    /// <summary>
    ///     Converts Array from GetBodyInversionMatrix() to double[].
    /// </summary>
    public static double[] GetBodyInversionMatrix(this Occurrence occurrence)
    {
        var matrix = Array.CreateInstance(typeof(double), 0);
        occurrence.GetBodyInversionMatrix(ref matrix);
        return matrix.OfType<double>().ToArray();
    }

    /// <summary>
    ///     Converts Array from GetExplodeMatrix() to double[].
    /// </summary>
    public static double[] GetExplodeMatrix(this Occurrence occurrence)
    {
        var matrix = Array.CreateInstance(typeof(double), 0);
        occurrence.GetExplodeMatrix(ref matrix);
        return matrix.OfType<double>().ToArray();
    }

    /// <summary>
    ///     Converts Array from GetMatrix() to double[].
    /// </summary>
    public static double[] GetMatrix(this Occurrence occurrence)
    {
        var matrix = Array.CreateInstance(typeof(double), 0);
        occurrence.GetMatrix(ref matrix);
        return matrix.OfType<double>().ToArray();
    }
}