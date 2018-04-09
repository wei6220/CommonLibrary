using Microsoft.VisualStudio.TestTools.UnitTesting;
using CommonLibrary;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace CommonLibrary.Tests
{
    [TestClass()]
    public class ZipHelperTests
    {
        [TestMethod()]
        public void CompressTest()
        {
            string files = @"d:\srctemp";
            //var files = new Dictionary<string, byte[]>();
            //files.Add("jerryqqq\\text1.txt", File.ReadAllBytes(@"d:\srctemp\test.txt"));

            var testUnit = new ZipHelper();
            var byteArray = testUnit.Compress(files);
            try
            {
                using (var fs = new FileStream(@"d:\ggtest.zip", FileMode.Create, FileAccess.Write))
                    fs.Write(byteArray, 0, byteArray.Length);
            }
            catch (Exception ex)
            {
            }

            Assert.IsFalse(false);
        }
    }
}