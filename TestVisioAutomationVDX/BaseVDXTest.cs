namespace TestVisioAutomationVDX
{
    using System;
    using System.IO;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using VisioAutomation.Extensions;
    using IVisio = Microsoft.Office.Interop.Visio;

    [TestClass]
    public class BaseVDXTest
    {
        private IVisio.Application application;
        private string output_directory;

        [TestInitialize]
        public void Initialize()
        {
            this.output_directory = Path.Combine(Path.GetTempPath(), "vdx-integration-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(this.output_directory);
        }

        [TestCleanup]
        public void Cleanup()
        {
            try
            {
                this.application?.Quit(true);
            }
            finally
            {
                this.application = null;
                Directory.Delete(this.output_directory, true);
            }
        }

        protected IVisio.Application GetApplication()
        {
            return this.application ?? (this.application = new IVisio.Application());
        }

        protected string NewOutputPath()
        {
            return Path.Combine(this.output_directory, Guid.NewGuid().ToString("N") + ".vdx");
        }

        protected string CopyForOpen(string filename)
        {
            string source = this.NewOutputPath();
            File.Copy(filename, source);
            return source;
        }

        protected string GetTestResultsOutPath(string path)
        {
            return Path.Combine(this.TestResultsOutFolder, path);
        }

        protected string TestResultsOutFolder => Path.GetDirectoryName(typeof(BaseVDXTest).Assembly.Location);
    }
}
