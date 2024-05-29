using SolidEdgePart;

namespace SolidEdgeCommunity.Extensions;

/// <summary>
///     SolidEdgePart.RefPlanes extension methods.
/// </summary>
public static class RefPlanesExtensions
{
    /// <summary>
    ///     Returns a SolidEdgePart.RefPlane representing the default top plane.
    /// </summary>
    public static RefPlane GetTopPlane(this RefPlanes refPlanes)
    {
        return refPlanes.Item(1);
    }

    /// <summary>
    ///     Returns a SolidEdgePart.RefPlane representing the default right plane.
    /// </summary>
    public static RefPlane GetRightPlane(this RefPlanes refPlanes)
    {
        return refPlanes.Item(2);
    }

    /// <summary>
    ///     Returns a SolidEdgePart.RefPlane representing the default front plane.
    /// </summary>
    public static RefPlane GetFrontPlane(this RefPlanes refPlanes)
    {
        return refPlanes.Item(3);
    }
}