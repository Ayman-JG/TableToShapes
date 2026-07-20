using System;
using System.Runtime.InteropServices;

namespace TableToShapes.AddIn
{
    // Hand-rolled COM interop definitions so we don't depend on the legacy
    // "Extensibility" PIA. GUIDs are fixed by Office and never change.

    /// <summary>Office add-in lifecycle interface (extensibility model).</summary>
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        void OnConnection([MarshalAs(UnmanagedType.IDispatch)] object application,
                          ext_ConnectMode connectMode,
                          [MarshalAs(UnmanagedType.IDispatch)] object addInInst,
                          ref Array custom);
        void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom);
        void OnAddInsUpdate(ref Array custom);
        void OnStartupComplete(ref Array custom);
        void OnBeginShutdown(ref Array custom);
    }

    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3,
        ext_cm_Solution = 4,
        ext_cm_UISetup = 5
    }

    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1,
        ext_dm_UISetupComplete = 2,
        ext_dm_SolutionClosed = 3
    }

    /// <summary>Lets the add-in supply Ribbon XML to Office.</summary>
    [ComImport]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonID);
    }

    /// <summary>Passed to Ribbon callbacks; identifies the control that fired.</summary>
    [ComImport]
    [Guid("000C0395-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonControl
    {
        string Id { get; }
        // Default VARIANT marshaling carries the IDispatch pointer; [return:]
        // attributes are not valid on properties.
        object Context { get; }
        string Tag { get; }
    }
}
