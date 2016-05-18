using System;
using ForecastToolMethods;
using NUnit.Framework;

namespace ExcelProc.Tests
{
    [TestFixture]
    public class ExcelProcTest
    {
        [Test]
        public void FinancialYearStartDateShouldBe26Mar2018ForYear2018()
        {
            DateTime exp = new DateTime(2018,3,26);
            var res = Methods.GetFinancialYearStartDate(new DateTime(2018,4,1));
            Assert.AreEqual(exp, res);
        }
    }
}
