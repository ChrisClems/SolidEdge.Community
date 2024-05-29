using System;
using System.Linq;
using SolidEdgePart;

namespace SolidEdgeCommunity.Extensions;

/// <summary>
///     SolidEdgePart.Arc3D extension methods.
/// </summary>
public static class Arc3DExtensions
{
    public static double[] SafeGetKeypointPosition(this Arc3D arc3d, Sketch3DKeypointType KeypointType,
        bool throwOnError = false)
    {
        var position = Array.CreateInstance(typeof(double), 0);

        try
        {
            // GetKeypointPosition may throw an exception so wrap in try\catch.
            arc3d.GetKeypointPosition(KeypointType, ref position);
        }
        catch
        {
            if (throwOnError) throw;
        }

        return position.OfType<double>().ToArray();
    }
}