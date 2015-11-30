using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MT940;

namespace UnitTestProjectMT940
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            Assert.AreEqual("01", MyMonth.MonthToDigit("января"));
            Assert.AreEqual("02", MyMonth.MonthToDigit("февраля"));
            Assert.AreEqual("03", MyMonth.MonthToDigit("марта"));
            Assert.AreEqual("04", MyMonth.MonthToDigit("апреля"));
            Assert.AreEqual("05", MyMonth.MonthToDigit("мая"));
            Assert.AreEqual("06", MyMonth.MonthToDigit("июня"));
            Assert.AreEqual("07", MyMonth.MonthToDigit("июля"));
            Assert.AreEqual("08", MyMonth.MonthToDigit("августа"));
            Assert.AreEqual("09", MyMonth.MonthToDigit("сентября"));
            Assert.AreEqual("10", MyMonth.MonthToDigit("октября"));
            Assert.AreEqual("11", MyMonth.MonthToDigit("ноября"));
            Assert.AreEqual("12", MyMonth.MonthToDigit("декабря"));
        }
    }
}
