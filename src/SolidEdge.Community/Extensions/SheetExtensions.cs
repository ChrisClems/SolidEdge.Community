using System;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using SolidEdgeDraft;

namespace SolidEdgeCommunity.Extensions;

public static class SheetExtensions
{
    private const uint CF_ENHMETAFILE = 14;

    [DllImport("user32.dll")]
    private static extern bool CloseClipboard();

    [DllImport("user32.dll", CharSet = CharSet.Unicode, ExactSpelling = true,
        CallingConvention = CallingConvention.StdCall)]
    private static extern IntPtr GetClipboardData(uint format);

    [DllImport("user32.dll")]
    private static extern IntPtr GetClipboardOwner();

    [DllImport("user32.dll")]
    private static extern bool IsClipboardFormatAvailable(uint format);

    [DllImport("user32.dll")]
    private static extern bool OpenClipboard(IntPtr hWndNewOwner);

    [DllImport("gdi32.dll")]
    private static extern bool DeleteEnhMetaFile(IntPtr hemf);

    [DllImport("gdi32.dll")]
    private static extern uint GetEnhMetaFileBits(IntPtr hemf, uint cbBuffer, [Out] byte[] lpbBuffer);

    public static Metafile GetEnhancedMetafile(this Sheet sheet)
    {
        try
        {
            // Copy the sheet as an EMF to the windows clipboard.
            sheet.CopyEMFToClipboard();

            if (OpenClipboard(IntPtr.Zero))
            {
                if (IsClipboardFormatAvailable(CF_ENHMETAFILE))
                {
                    // Get the handle to the EMF.
                    var henhmetafile = GetClipboardData(CF_ENHMETAFILE);

                    return new Metafile(henhmetafile, true);
                }
                else
                {
                    throw new Exception("CF_ENHMETAFILE is not available in clipboard.");
                }
            }
            else
            {
                throw new Exception("Error opening clipboard.");
            }
        }
        finally
        {
            CloseClipboard();
        }
    }

    public static void SaveAsEnhancedMetafile(this Sheet sheet, string filename)
    {
        try
        {
            // Copy the sheet as an EMF to the windows clipboard.
            sheet.CopyEMFToClipboard();

            if (OpenClipboard(IntPtr.Zero))
            {
                if (IsClipboardFormatAvailable(CF_ENHMETAFILE))
                {
                    // Get the handle to the EMF.
                    var hEMF = GetClipboardData(CF_ENHMETAFILE);

                    // Query the size of the EMF.
                    var len = GetEnhMetaFileBits(hEMF, 0, null);
                    var rawBytes = new byte[len];

                    // Get all of the bytes of the EMF.
                    GetEnhMetaFileBits(hEMF, len, rawBytes);

                    // Write all of the bytes to a file.
                    File.WriteAllBytes(filename, rawBytes);

                    // Delete the EMF handle.
                    DeleteEnhMetaFile(hEMF);
                }
                else
                {
                    throw new Exception("CF_ENHMETAFILE is not available in clipboard.");
                }
            }
            else
            {
                throw new Exception("Error opening clipboard.");
            }
        }
        finally
        {
            CloseClipboard();
        }
    }
}