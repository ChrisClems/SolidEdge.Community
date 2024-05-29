﻿using System;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using SolidEdgeConstants;
using SolidEdgeFramework;
using Application = SolidEdgeFramework.Application;
using ApplicationGlobalConstants = SolidEdgeFramework.ApplicationGlobalConstants;
using Environment = SolidEdgeFramework.Environment;
using SolidEdgeCommandConstants = SolidEdgeConstants.SolidEdgeCommandConstants;

namespace SolidEdgeCommunity.Extensions;

/// <summary>
///     SolidEdgeFramework.Application extension methods.
/// </summary>
public static class ApplicationExtensions
{
    /// <summary>
    ///     Creates a Solid Edge command control.
    /// </summary>
    /// <param name="application"></param>
    /// <param name="CmdFlags"></param>
    public static Command CreateCommand(this Application application, seCmdFlag CmdFlags)
    {
        return application.CreateCommand((int)CmdFlags);
    }

    /// <summary>
    ///     Returns the currently active document.
    /// </summary>
    /// <remarks>An exception will be thrown if there is no active document.</remarks>
    public static SolidEdgeDocument GetActiveDocument(this Application application)
    {
        return application.GetActiveDocument(true);
    }

    /// <summary>
    ///     Returns the currently active document.
    /// </summary>
    public static SolidEdgeDocument GetActiveDocument(this Application application, bool throwOnError)
    {
        try
        {
            // ActiveDocument will throw an exception if no document is open.
            return (SolidEdgeDocument)application.ActiveDocument;
        }
        catch
        {
            if (throwOnError) throw;
        }

        return null;
    }

    /// <summary>
    ///     Returns the currently active document.
    /// </summary>
    /// <typeparam name="T">The type to return.</typeparam>
    /// ///
    /// <remarks>An exception will be thrown if there is no active document or if the cast fails.</remarks>
    public static T GetActiveDocument<T>(this Application application) where T : class
    {
        return application.GetActiveDocument<T>(true);
    }

    /// <summary>
    ///     Returns the currently active document.
    /// </summary>
    /// <typeparam name="T">The type to return.</typeparam>
    public static T GetActiveDocument<T>(this Application application, bool throwOnError) where T : class
    {
        try
        {
            // ActiveDocument will throw an exception if no document is open.
            return (T)application.ActiveDocument;
        }
        catch
        {
            if (throwOnError) throw;
        }

        return null;
    }

    /// <summary>
    ///     Returns the environment that belongs to the current active window of the document.
    /// </summary>
    public static Environment GetActiveEnvironment(this Application application)
    {
        var environments = application.Environments;
        return environments.Item(application.ActiveEnvironment);
    }

    ///// <summary>
    ///// Returns the application events.
    ///// </summary>
    //public static SolidEdgeFramework.ISEApplicationEvents_Event GetApplicationEvents(this SolidEdgeFramework.Application application)
    //{
    //    return (SolidEdgeFramework.ISEApplicationEvents_Event)application.ApplicationEvents;
    //}

    ///// <summary>
    ///// Returns the application window events.
    ///// </summary>
    //public static SolidEdgeFramework.ISEApplicationWindowEvents_Event GetApplicationWindowEvents(this SolidEdgeFramework.Application application)
    //{
    //    return (SolidEdgeFramework.ISEApplicationWindowEvents_Event)application.ApplicationWindowEvents;
    //}

    /// <summary>
    ///     Returns an environment specified by CATID.
    /// </summary>
    public static Environment GetEnvironment(this Application application, string CATID)
    {
        var guid1 = new Guid(CATID);

        foreach (var environment in application.Environments.OfType<Environment>())
        {
            var guid2 = new Guid(environment.CATID);
            if (guid1.Equals(guid2)) return environment;
        }

        return null;
    }

    ///// <summary>
    ///// Returns the feature library events.
    ///// </summary>
    //public static SolidEdgeFramework.ISEFeatureLibraryEvents_Event GetFeatureLibraryEvents(this SolidEdgeFramework.Application application)
    //{
    //    return (SolidEdgeFramework.ISEFeatureLibraryEvents_Event)application.FeatureLibraryEvents;
    //}

    ///// <summary>
    ///// Returns the file UI events.
    ///// </summary>
    //public static SolidEdgeFramework.ISEFileUIEvents_Event GetFileUIEvents(this SolidEdgeFramework.Application application)
    //{
    //    return (SolidEdgeFramework.ISEFileUIEvents_Event)application.FileUIEvents;
    //}

    /// <summary>
    ///     Returns the value of a specified global constant.
    /// </summary>
    public static object GetGlobalParameter(this Application application, ApplicationGlobalConstants globalConstant)
    {
        object value = null;
        application.GetGlobalParameter(globalConstant, ref value);
        return value;
    }

    /// <summary>
    ///     Returns the value of a specified global constant.
    /// </summary>
    /// <typeparam name="T">The type to return.</typeparam>
    public static T GetGlobalParameter<T>(this Application application, ApplicationGlobalConstants globalConstant)
    {
        object value = null;
        application.GetGlobalParameter(globalConstant, ref value);
        return (T)value;
    }

    /// <summary>
    ///     Returns a NativeWindow object that represents the main application window.
    /// </summary>
    public static NativeWindow GetNativeWindow(this Application application)
    {
        return NativeWindow.FromHandle(new IntPtr(application.hWnd));
    }

    ///// <summary>
    ///// Returns the new file UI events.
    ///// </summary>
    //public static SolidEdgeFramework.ISENewFileUIEvents_Event GetNewFileUIEvents(this SolidEdgeFramework.Application application)
    //{
    //    return (SolidEdgeFramework.ISENewFileUIEvents_Event)application.NewFileUIEvents;
    //}

    ///// <summary>
    ///// Returns the SEEC events.
    ///// </summary>
    //public static SolidEdgeFramework.ISEECEvents_Event GetSEECEvents(this SolidEdgeFramework.Application application)
    //{
    //    return (SolidEdgeFramework.ISEECEvents_Event)application.SEECEvents;
    //}

    /// <summary>
    ///     Returns a Process object that represents the application prcoess.
    /// </summary>
    public static Process GetProcess(this Application application)
    {
        return Process.GetProcessById(application.ProcessID);
    }

    ///// <summary>
    ///// Returns the shortcut menu events.
    ///// </summary>
    //public static SolidEdgeFramework.ISEShortCutMenuEvents_Event GetShortcutMenuEvents(this SolidEdgeFramework.Application application)
    //{
    //    return (SolidEdgeFramework.ISEShortCutMenuEvents_Event)application.ShortcutMenuEvents;
    //}

    /// <summary>
    ///     Returns a Version object that represents the application version.
    /// </summary>
    public static Version GetVersion(this Application application)
    {
        return new Version(application.Version);
    }

    /// <summary>
    ///     Returns a point object to the application main window.
    /// </summary>
    public static IntPtr GetWindowHandle(this Application application)
    {
        return new IntPtr(application.hWnd);
    }

    /// <summary>
    ///     Activates a specified Solid Edge command.
    /// </summary>
    public static void StartCommand(this Application application, AssemblyCommandConstants CommandID)
    {
        application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
    }

    /// <summary>
    ///     Activates a specified Solid Edge command.
    /// </summary>
    public static void StartCommand(this Application application, CuttingPlaneLineCommandConstants CommandID)
    {
        application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
    }

    /// <summary>
    ///     Activates a specified Solid Edge command.
    /// </summary>
    public static void StartCommand(this Application application, DetailCommandConstants CommandID)
    {
        application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
    }

    /// <summary>
    ///     Activates a specified Solid Edge command.
    /// </summary>
    public static void StartCommand(this Application application, DrawingViewEditCommandConstants CommandID)
    {
        application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
    }

    /// <summary>
    ///     Activates a specified Solid Edge command.
    /// </summary>
    public static void StartCommand(this Application application, ExplodeCommandConstants CommandID)
    {
        application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
    }

    /// <summary>
    ///     Activates a specified Solid Edge command.
    /// </summary>
    public static void StartCommand(this Application application, LayoutCommandConstants CommandID)
    {
        application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
    }

    /// <summary>
    ///     Activates a specified Solid Edge command.
    /// </summary>
    public static void StartCommand(this Application application, LayoutInPartCommandConstants CommandID)
    {
        application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
    }

    /// <summary>
    ///     Activates a specified Solid Edge command.
    /// </summary>
    public static void StartCommand(this Application application, MotionCommandConstants CommandID)
    {
        application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
    }

    /// <summary>
    ///     Activates a specified Solid Edge command.
    /// </summary>
    public static void StartCommand(this Application application, PartCommandConstants CommandID)
    {
        application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
    }

    /// <summary>
    ///     Activates a specified Solid Edge command.
    /// </summary>
    public static void StartCommand(this Application application, ProfileCommandConstants CommandID)
    {
        application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
    }

    /// <summary>
    ///     Activates a specified Solid Edge command.
    /// </summary>
    public static void StartCommand(this Application application, ProfileHoleCommandConstants CommandID)
    {
        application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
    }

    /// <summary>
    ///     Activates a specified Solid Edge command.
    /// </summary>
    public static void StartCommand(this Application application, ProfilePatternCommandConstants CommandID)
    {
        application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
    }

    /// <summary>
    ///     Activates a specified Solid Edge command.
    /// </summary>
    public static void StartCommand(this Application application, ProfileRevolvedCommandConstants CommandID)
    {
        application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
    }

    /// <summary>
    ///     Activates a specified Solid Edge command.
    /// </summary>
    public static void StartCommand(this Application application, SheetMetalCommandConstants CommandID)
    {
        application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
    }

    /// <summary>
    ///     Activates a specified Solid Edge command.
    /// </summary>
    public static void StartCommand(this Application application, SimplifyCommandConstants CommandID)
    {
        application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
    }

    /// <summary>
    ///     Activates a specified Solid Edge command.
    /// </summary>
    public static void StartCommand(this Application application, SolidEdgeCommandConstants CommandID)
    {
        application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
    }

    /// <summary>
    ///     Activates a specified Solid Edge command.
    /// </summary>
    public static void StartCommand(this Application application, StudioCommandConstants CommandID)
    {
        application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
    }

    /// <summary>
    ///     Activates a specified Solid Edge command.
    /// </summary>
    public static void StartCommand(this Application application, TubingCommandConstants CommandID)
    {
        application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
    }

    /// <summary>
    ///     Activates a specified Solid Edge command.
    /// </summary>
    public static void StartCommand(this Application application, WeldmentCommandConstants CommandID)
    {
        application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
    }

    /// <summary>
    ///     Activates a specified Solid Edge command.
    /// </summary>
    public static void StartCommand(this Application application,
        SolidEdgeFramework.SolidEdgeCommandConstants CommandID)
    {
        application.StartCommand(CommandID);
    }

    /// <summary>
    ///     Shows the form with the application main window as the owner.
    /// </summary>
    public static void Show(this Application application, Form form)
    {
        if (form == null) throw new ArgumentNullException("form");

        form.Show(application.GetNativeWindow());
    }

    /// <summary>
    ///     Shows the form as a modal dialog box with the application main window as the owner.
    /// </summary>
    public static DialogResult ShowDialog(this Application application, Form form)
    {
        if (form == null) throw new ArgumentNullException("form");

        return form.ShowDialog(application.GetNativeWindow());
    }

    /// <summary>
    ///     Shows the form as a modal dialog box with the application main window as the owner.
    /// </summary>
    public static DialogResult ShowDialog(this Application application, CommonDialog dialog)
    {
        if (dialog == null) throw new ArgumentNullException("dialog");
        return dialog.ShowDialog(application.GetNativeWindow());
    }
}