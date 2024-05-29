using System;
using SolidEdgeFramework;

namespace SolidEdgeCommunity.Extensions;

/// <summary>
///     SolidEdgeFramework.Window extension methods.
/// </summary>
public static class WindowExtensions
{
    /// <summary>
    ///     Returns an IntPtr representing the window handle.
    /// </summary>
    public static IntPtr GetDrawHandle(this Window window)
    {
        return new IntPtr(window.DrawHwnd);
    }

    /// <summary>
    ///     Returns an IntPtr representing the window handle.
    /// </summary>
    public static IntPtr GetHandle(this Window window)
    {
        return new IntPtr(window.hWnd);
    }
}