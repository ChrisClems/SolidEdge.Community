using System;
using SolidEdgeConstants;
using SolidEdgeFramework;
using SolidEdgeSDK;
using Environment = SolidEdgeFramework.Environment;
using SolidEdgeCommandConstants = SolidEdgeConstants.SolidEdgeCommandConstants;

namespace SolidEdgeCommunity.Extensions;

/// <summary>
///     SolidEdgeFramework.Environment extension methods.
/// </summary>
public static class EnvironmentExtensions
{
    /// <summary>
    ///     Returns a Guid representing the environment category.
    /// </summary>
    public static Guid GetCategoryId(this Environment environment)
    {
        return new Guid(environment.CATID);
    }

    /// <summary>
    ///     Returns a Type representing the corresponding command constants for a particular environment.
    /// </summary>
    /// <param name="environment"></param>
    /// <returns></returns>
    public static Type GetCommandConstantType(this Environment environment)
    {
        var categoryId = environment.GetCategoryId();

        if (categoryId.Equals(EnvironmentCategories.Application))
            return typeof(SolidEdgeCommandConstants);
        if (categoryId.Equals(EnvironmentCategories.Assembly))
            return typeof(AssemblyCommandConstants);
        if (categoryId.Equals(EnvironmentCategories.DMAssembly))
            return typeof(AssemblyCommandConstants);
        if (categoryId.Equals(EnvironmentCategories.CuttingPlaneLine))
            return typeof(CuttingPlaneLineCommandConstants);
        if (categoryId.Equals(EnvironmentCategories.Draft))
            return typeof(DetailCommandConstants);
        if (categoryId.Equals(EnvironmentCategories.DrawingViewEdit))
            return typeof(DrawingViewEditCommandConstants);
        if (categoryId.Equals(EnvironmentCategories.Explode))
            return typeof(ExplodeCommandConstants);
        if (categoryId.Equals(EnvironmentCategories.Layout))
            return typeof(LayoutCommandConstants);
        if (categoryId.Equals(EnvironmentCategories.Sketch))
            return typeof(LayoutInPartCommandConstants);
        if (categoryId.Equals(EnvironmentCategories.Motion))
            return typeof(MotionCommandConstants);
        if (categoryId.Equals(EnvironmentCategories.Part))
            return typeof(PartCommandConstants);
        if (categoryId.Equals(EnvironmentCategories.DMPart))
            return typeof(PartCommandConstants);
        if (categoryId.Equals(EnvironmentCategories.Profile))
            return typeof(ProfileCommandConstants);
        if (categoryId.Equals(EnvironmentCategories.ProfileHole))
            return typeof(ProfileHoleCommandConstants);
        if (categoryId.Equals(EnvironmentCategories.ProfilePattern))
            return typeof(ProfilePatternCommandConstants);
        if (categoryId.Equals(EnvironmentCategories.ProfileRevolved))
            return typeof(ProfileRevolvedCommandConstants);
        if (categoryId.Equals(EnvironmentCategories.SheetMetal))
            return typeof(SheetMetalCommandConstants);
        if (categoryId.Equals(EnvironmentCategories.DMSheetMetal))
            return typeof(SheetMetalCommandConstants);
        if (categoryId.Equals(EnvironmentCategories.Simplify))
            return typeof(SimplifyCommandConstants);
        if (categoryId.Equals(EnvironmentCategories.Studio))
            return typeof(StudioCommandConstants);
        if (categoryId.Equals(EnvironmentCategories.XpresRoute))
            return typeof(TubingCommandConstants);
        if (categoryId.Equals(EnvironmentCategories.Weldment)) return typeof(WeldmentCommandConstants);

        return null;
    }

    /// <summary>
    ///     Returns a SolidEdgeFramework.Environment by name.
    /// </summary>
    public static Environment LookupByName(this Environments environments, string name)
    {
        for (var i = 1; i <= environments.Count; i++)
        {
            var environment = environments.Item(i);
            if (environment.Name.Equals(name)) return environment;
        }

        return null;
    }

    /// <summary>
    ///     Returns a SolidEdgeFramework.Environment by category id..
    /// </summary>
    public static Environment LookupByCategoryId(this Environments environments, Guid categoryId)
    {
        for (var i = 1; i <= environments.Count; i++)
        {
            var environment = environments.Item(i);
            if (environment.GetCategoryId().Equals(categoryId)) return environment;
        }

        return null;
    }
}