using Xunit;
using System.Runtime.InteropServices;
using SolidEdgeCommunity;

namespace SolidEdge.Community.Tests;

public class ConnectTest
{
    [Fact]
    public void StartAndConnectToSolidEdge()
    {
        // Arrange
        const bool startIfNotRunning = true;
        const bool ensureVisible = true;
        
        // Act
        var app = SolidEdgeUtils.Connect(startIfNotRunning, ensureVisible);

        // Assert
        Assert.NotNull(app);
        Assert.True(app is SolidEdgeFramework.Application); //check if Solid Edge COM object is available
    }
}