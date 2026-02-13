using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace TasteDocumentGenerator
{
    public static class DeploymentViewHelper
    {
        /// <summary>
        /// Returns the name of the partition that contains the largest number of functions in the
        /// provided Deployment View (DV) XML file. Returns null if the file doesn't exist or no
        /// suitable partition/name can be determined.
        /// </summary>
        public static string? GetTargetName(string deploymentViewPath)
        {
            if (string.IsNullOrWhiteSpace(deploymentViewPath) || !File.Exists(deploymentViewPath))
                return null;

            try
            {
                var doc = XDocument.Load(deploymentViewPath);

                // Select Partition elements regardless of namespace
                var partitions = doc.Descendants().Where(e => string.Equals(e.Name.LocalName, "Partition", StringComparison.OrdinalIgnoreCase));
                if (!partitions.Any())
                    return null;

                string? bestName = null;
                int bestCount = -1;

                foreach (var p in partitions)
                {
                    // Count Function elements under this partition (search descendants to be tolerant)
                    var functionCount = p.Descendants().Count(e => string.Equals(e.Name.LocalName, "Function", StringComparison.OrdinalIgnoreCase));
                    if (functionCount > bestCount)
                    {
                        bestCount = functionCount;
                        bestName = p.Attribute("name")?.Value;
                    }
                }

                return bestName;
            }
            catch
            {
                return null;
            }
        }
    }
}
