using System.Collections.Generic;
using SolidEdgeDraft;

namespace SolidEdgeCommunity.Extensions;

/// <summary>
///     SolidEdgeDraft.Section extensions.
/// </summary>
public static class SectionExtensions
{
    /// <summary>
    ///     Returns an enumerable collection of drawing objects.
    /// </summary>
    /// <param name="section"></param>
    /// <returns></returns>
    public static IEnumerable<object> EnumerateDrawingObjects(this Section section)
    {
        foreach (var drawingObject in EnumerateDrawingObjects<object>(section)) yield return drawingObject;
    }

    /// <summary>
    ///     Returns an enumerable collection of drawing objects of the specified type.
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    public static IEnumerable<T> EnumerateDrawingObjects<T>(this Section section) where T : class
    {
        foreach (Sheet sheet in section.Sheets)
        {
            // The following section types are hidden and should not be used.
            if (sheet.SectionType == SheetSectionTypeConstants.igDrawingViewSection) continue;
            if (sheet.SectionType == SheetSectionTypeConstants.igBlockViewSection) continue;

            // Should work but throws an exception...
            //foreach (var drawingObject in sheet.DrawingObjects)

            var drawingObjects = sheet.DrawingObjects;

            for (var i = 1; i <= drawingObjects.Count; i++)
            {
                var drawingObject = drawingObjects.Item(i);

                if (drawingObject is T) yield return (T)drawingObject;
            }
        }
    }

    /// <summary>
    ///     Returns an enumerable collection of drawing views.
    /// </summary>
    /// <param name="section"></param>
    /// <returns></returns>
    public static IEnumerable<DrawingView> EnumerateDrawingViews(this Section section)
    {
        foreach (Sheet sheet in section.Sheets)
        foreach (DrawingView drawingView in sheet.DrawingViews)
            yield return drawingView;
    }
}