using System;
using System.Collections.Generic;
using SolidEdgeDraft;
using SolidEdgeFramework;

namespace SolidEdgeCommunity.Extensions;

/// <summary>
///     SolidEdgeDraft.DraftDocument extensions.
/// </summary>
public static class DraftDocumentExtensions
{
    /// <summary>
    ///     Returns an enumerable collection of drawing objects.
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    public static IEnumerable<object> EnumerateDrawingObjects(this DraftDocument document)
    {
        foreach (Section section in document.Sections)
        foreach (var drawingObject in section.EnumerateDrawingObjects())
            yield return drawingObject;
    }

    /// <summary>
    ///     Returns an enumerable collection of drawing objects of the specified type.
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    public static IEnumerable<T> EnumerateDrawingObjects<T>(this DraftDocument document) where T : class
    {
        foreach (Section section in document.Sections)
        foreach (var drawingObject in section.EnumerateDrawingObjects<T>())
            yield return drawingObject;
    }

    /// <summary>
    ///     Returns an enumerable collection of drawing objects in the specified section.
    /// </summary>
    /// <param name="document"></param>
    /// <param name="sectionType"></param>
    /// <returns></returns>
    public static IEnumerable<object> EnumerateDrawingObjects(this DraftDocument document,
        SheetSectionTypeConstants sectionType)
    {
        foreach (Section section in document.Sections)
            if (section.Type == sectionType)
                foreach (var drawingObject in section.EnumerateDrawingObjects())
                    yield return drawingObject;
    }

    /// <summary>
    ///     Returns an enumerable collection of drawing objects of the specified type in the specified section.
    /// </summary>
    /// <param name="document"></param>
    /// <param name="sectionType"></param>
    /// <returns></returns>
    public static IEnumerable<object> EnumerateDrawingObjects<T>(this DraftDocument document,
        SheetSectionTypeConstants sectionType) where T : class
    {
        foreach (Section section in document.Sections)
            if (section.Type == sectionType)
                foreach (var drawingObject in section.EnumerateDrawingObjects<T>())
                    yield return drawingObject;
    }

    /// <summary>
    ///     Returns an enumerable collection of drawing views.
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    public static IEnumerable<DrawingView> EnumerateDrawingViews(this DraftDocument document)
    {
        foreach (Sheet sheet in document.Sheets)
        foreach (DrawingView drawingView in sheet.DrawingViews)
            yield return drawingView;
    }

    /// <summary>
    ///     Returns an enumerable collection of drawing views in the specified section.
    /// </summary>
    /// <param name="document"></param>
    /// <param name="sectionType"></param>
    /// <returns></returns>
    public static IEnumerable<DrawingView> EnumerateDrawingViews(this DraftDocument document,
        SheetSectionTypeConstants sectionType)
    {
        foreach (Section section in document.Sections)
            if (section.Type == sectionType)
                foreach (var drawingView in section.EnumerateDrawingViews())
                    yield return drawingView;
    }

    /// <summary>
    ///     Returns the version of Solid Edge that was used to create the referenced document.
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    public static Version GetCreatedVersion(this DraftDocument document)
    {
        return new Version(document.CreatedVersion);
    }

    /// <summary>
    ///     Returns the version of Solid Edge that was used the last time the referenced document was saved.
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    public static Version GetLastSavedVersion(this DraftDocument document)
    {
        return new Version(document.LastSavedVersion);
    }

    /// <summary>
    ///     Returns the properties for the referenced document.
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    public static PropertySets GetProperties(this DraftDocument document)
    {
        return document.Properties as PropertySets;
    }

    /// <summary>
    ///     Returns the summary information property set for the referenced document.
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    public static SummaryInfo GetSummaryInfo(this DraftDocument document)
    {
        return document.SummaryInfo as SummaryInfo;
    }

    /// <summary>
    ///     Returns a collection of variables for the referenced document.
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    public static Variables GetVariables(this DraftDocument document)
    {
        return document.Variables as Variables;
    }
}