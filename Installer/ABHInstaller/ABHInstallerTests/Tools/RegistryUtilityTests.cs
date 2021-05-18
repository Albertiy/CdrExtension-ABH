using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Win32;
using System;
using System.Collections;

namespace ABHInstaller.Tools.Tests
{
    [TestClass()]
    public class RegistryUtilityTests
    {
        [TestMethod()]
        public void GetRegistryValueTest()
        {
            string res = RegistryUtility.GetRegistryValue(Registry.LocalMachine, @"SOFTWARE\Corel\Setup\CorelDRAW Graphics Suite 17", "Destination");
            Console.WriteLine("res: {0}", res);
            //Assert.Fail();
        }

        [TestMethod()]
        public void GetRegistryItemsTest()
        {
            ArrayList res = RegistryUtility.GetRegistryItems(Registry.LocalMachine, @"SOFTWARE\WOW6432Node\Corel\Setup");
            Console.WriteLine("res: {0}", string.Join(", ", (string[])res.ToArray(typeof(string))));
        }
    }
}