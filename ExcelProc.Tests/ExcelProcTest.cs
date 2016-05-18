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

        [Test]
        public void BillingPeriodShouldBe29Jun2020To24July2020WhenDateIs1July2020()
        {
            DateTime[] expDateTimes = new DateTime[2];
            expDateTimes[0] = new DateTime(2020, 06, 29);
            expDateTimes[1] = new DateTime(2020, 07, 24);
            var res = Methods.GetBillingPeriodGeneral(new DateTime(2020, 07, 01));
            Assert.AreEqual(expDateTimes, res);
        }
    }

    [TestFixture]
    public class DateRangeTest
    {
        [Test]
        public void Range25May2020To30Jun2020ShouldIncludeDate15Jun2020()
        {
            var dr = new DateRange(new DateTime(2020, 05, 25), new DateTime(2020, 06, 30));
            var testDate = new DateTime(2020, 06, 15);
            Assert.IsTrue(dr.Includes(testDate));
        }

        [Test]
        public void Range25May2020To30Jun2020ShouldIncludeDatesBetween01Jun2020To15Jun2020()
        {
            var dr = new DateRange(new DateTime(2020, 05, 25), new DateTime(2020, 06, 30));
            var testDate = new DateRange(new DateTime(2020, 06, 01), new DateTime(2020, 06, 15));
            Assert.IsTrue(dr.Includes(testDate));
        }
    }
}
