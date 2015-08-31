using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using BassUtils;

namespace TestParser.Core
{
    public class NCrunchCoverageParser
    {
        public IEnumerable<CoverageData> Parse(string filename)
        {
            filename.ThrowIfFileDoesNotExist("filename");

            var doc = XDocument.Load(filename);

            try
            {
                var coverage = (from proj in doc.Descendants("project")
                                let projPath = proj.Attribute("path").Value
                                from src in proj.Descendants("sourceFile")
                                let srcPath = src.Attribute("path").Value
                                //from line in src.Descendants("line")
                                select new CoverageData()
                                {
                                    ProjectPathName = projPath,
                                    SourceFilePathName = srcPath,
                                    CompiledLines = src.Descendants("line").Count(),
                                    CoveredLines = src.Descendants("line").Where(x => x.Attribute("coveringTests").Value != "0").Count()
                                }
                                ).OrderBy(c => c.ProjectPathName)
                                 .ThenBy(c => c.SourceFilePathName);

                return coverage;
            }
            catch (Exception ex)
            {
                throw new ParseException("Error while parsing NUnit file '" + filename + "'", ex);
            }

        }
    }
}