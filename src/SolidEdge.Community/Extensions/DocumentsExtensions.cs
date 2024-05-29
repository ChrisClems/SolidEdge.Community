using System;
using SolidEdgeAssembly;
using SolidEdgeDraft;
using SolidEdgeFramework;
using SolidEdgePart;
using SolidEdgeSDK;

namespace SolidEdgeCommunity.Extensions;

/// <summary>
///     SolidEdgeFramework.Documents extension methods.
/// </summary>
public static class DocumentsExtensions
{
    /// <summary>
    ///     Creates a new document.
    /// </summary>
    internal static TDocumentType Add<TDocumentType>(this Documents documents, string progId)
        where TDocumentType : class
    {
        return (TDocumentType)documents.Add(progId);
    }

    /// <summary>
    ///     Creates a new document.
    /// </summary>
    internal static TDocumentType Add<TDocumentType>(this Documents documents, string progId, object TemplateDoc)
        where TDocumentType : class
    {
        return (TDocumentType)documents.Add(progId, TemplateDoc);
    }

    /// <summary>
    ///     Creates a new document.
    /// </summary>
    /// <typeparam name="TDocumentType">
    ///     SolidEdgeAssembly.AssemblyDocument, SolidEdgeDraft.DraftDocument,
    ///     SolidEdgePart.PartDocument or SolidEdgePart.SheetMetalDocument.
    /// </typeparam>
    /// <param name="documents"></param>
    /// <returns></returns>
    public static TDocumentType Add<TDocumentType>(this Documents documents) where TDocumentType : class
    {
        var t = typeof(TDocumentType);

        if (t.Equals(typeof(AssemblyDocument)))
            return (TDocumentType)documents.AddAssemblyDocument();
        if (t.Equals(typeof(DraftDocument)))
            return (TDocumentType)documents.AddDraftDocument();
        if (t.Equals(typeof(PartDocument)))
            return (TDocumentType)documents.AddPartDocument();
        if (t.Equals(typeof(SheetMetalDocument)))
            return (TDocumentType)documents.AddSheetMetalDocument();
        throw new NotSupportedException();
    }

    /// <summary>
    ///     Creates a new assembly document.
    /// </summary>
    public static AssemblyDocument AddAssemblyDocument(this Documents documents)
    {
        return documents.Add<AssemblyDocument>(PROGID.SolidEdge_AssemblyDocument);
    }

    /// <summary>
    ///     Creates a new assembly document.
    /// </summary>
    public static AssemblyDocument AddAssemblyDocument(this Documents documents, object TemplateDoc)
    {
        return documents.Add<AssemblyDocument>(PROGID.SolidEdge_AssemblyDocument, TemplateDoc);
    }

    /// <summary>
    ///     Creates a new draft document.
    /// </summary>
    public static DraftDocument AddDraftDocument(this Documents documents)
    {
        return documents.Add<DraftDocument>(PROGID.SolidEdge_DraftDocument);
    }

    /// <summary>
    ///     Creates a new draft document.
    /// </summary>
    public static DraftDocument AddDraftDocument(this Documents documents, object TemplateDoc)
    {
        return documents.Add<DraftDocument>(PROGID.SolidEdge_DraftDocument, TemplateDoc);
    }

    /// <summary>
    ///     Creates a new part document.
    /// </summary>
    public static PartDocument AddPartDocument(this Documents documents)
    {
        return documents.Add<PartDocument>(PROGID.SolidEdge_PartDocument);
    }

    /// <summary>
    ///     Creates a new part document.
    /// </summary>
    public static PartDocument AddPartDocument(this Documents documents, object TemplateDoc)
    {
        return documents.Add<PartDocument>(PROGID.SolidEdge_PartDocument, TemplateDoc);
    }

    /// <summary>
    ///     Creates a new sheetmetal document.
    /// </summary>
    public static SheetMetalDocument AddSheetMetalDocument(this Documents documents)
    {
        return documents.Add<SheetMetalDocument>(PROGID.SolidEdge_SheetMetalDocument);
    }

    /// <summary>
    ///     Creates a new sheetmetal document.
    /// </summary>
    public static SheetMetalDocument AddSheetMetalDocument(this Documents documents, object TemplateDoc)
    {
        return documents.Add<SheetMetalDocument>(PROGID.SolidEdge_SheetMetalDocument, TemplateDoc);
    }

    /// <summary>
    ///     Opens an existing Solid Edge document.
    /// </summary>
    public static TDocumentType Open<TDocumentType>(this Documents documents, string Filename)
        where TDocumentType : class
    {
        return (TDocumentType)documents.Open(Filename);
    }

    /// <summary>
    ///     Opens an existing Solid Edge document.
    /// </summary>
    public static TDocumentType Open<TDocumentType>(this Documents documents, string Filename,
        object DocRelationAutoServer, object AltPath, object RecognizeFeaturesIfPartTemplate, object RevisionRuleOption,
        object StopFileOpenIfRevisionRuleNotApplicable)
    {
        return (TDocumentType)documents.Open(Filename, DocRelationAutoServer, AltPath, RecognizeFeaturesIfPartTemplate,
            RevisionRuleOption, StopFileOpenIfRevisionRuleNotApplicable);
    }

    /// <summary>
    ///     Opens an existing Solid Edge document in the background with no window.
    /// </summary>
    public static TDocumentType OpenInBackground<TDocumentType>(this Documents documents, string Filename)
    {
        ulong JDOCUMENTPROP_NOWINDOW = 0x00000008;
        return (TDocumentType)documents.Open(Filename, JDOCUMENTPROP_NOWINDOW);
    }

    /// <summary>
    ///     Opens an existing Solid Edge document in the background with no window.
    /// </summary>
    public static TDocumentType OpenInBackground<TDocumentType>(this Documents documents, string Filename,
        object AltPath, object RecognizeFeaturesIfPartTemplate, object RevisionRuleOption,
        object StopFileOpenIfRevisionRuleNotApplicable)
    {
        ulong JDOCUMENTPROP_NOWINDOW = 0x00000008;
        return (TDocumentType)documents.Open(Filename, JDOCUMENTPROP_NOWINDOW, AltPath, RecognizeFeaturesIfPartTemplate,
            RevisionRuleOption, StopFileOpenIfRevisionRuleNotApplicable);
    }
}