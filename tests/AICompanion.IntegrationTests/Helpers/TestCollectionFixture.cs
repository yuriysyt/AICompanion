using Xunit;

// All [Collection("RealApp")] tests share the Windows desktop as a resource.
// Disable parallelisation globally so they never run at the same time.
[assembly: CollectionBehavior(DisableTestParallelization = true)]

namespace AICompanion.IntegrationTests.Helpers
{
    /// <summary>
    /// Marker definition for the "RealApp" xUnit collection.
    /// Tests decorated with [Collection("RealApp")] run sequentially in the
    /// order the runner discovers them, preventing desktop races.
    /// </summary>
    [CollectionDefinition("RealApp")]
    public class RealAppCollection { }
}
