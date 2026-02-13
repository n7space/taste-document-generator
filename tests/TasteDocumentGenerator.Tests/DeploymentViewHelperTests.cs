namespace TasteDocumentGenerator.Tests;

public class DeploymentViewHelperTests
{
    [Fact]
    public void GetTargetName_ReturnsPartitionWithMostFunctions()
    {
        var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempDir);
        var dvPath = Path.Combine(tempDir, "dv.xml");

        var xml = @"<?xml version=""1.0""?>
<DeploymentView>
  <Partition id=""{p1}"" name=""PartitionA"">
    <Function id=""{f1}"" name=""F1"" />
    <Function id=""{f2}"" name=""F2"" />
  </Partition>
  <Partition id=""p2"" name=""PartitionB"">
    <Function id=""{f3}"" name=""F3"" />
  </Partition>
</DeploymentView>";

        File.WriteAllText(dvPath, xml);

        try
        {
            var result = DeploymentViewHelper.GetTargetName(dvPath);
            Assert.Equal("PartitionA", result);
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }
}
