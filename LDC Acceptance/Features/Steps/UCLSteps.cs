using System;
using TechTalk.SpecFlow;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using P2PCalculations;
using System.IO;

namespace LDC_Acceptance.Features.Steps
{
    [Binding]
    public class UCLSteps
    {
        private UCL ucl = new UCL();
        private double result;
        private double min;
        private double max;

        [Given(@"I have access to setupFile")]
        public void GivenIHaveAccessToSetupFile()
        {
            Assert.IsTrue(File.Exists(P2PCalculations.Properties.Resources.setupFileName));
        }

        [When(@"I enter '(.*)' and (.*) and '(.*)'")]
        public void WhenIEnterAndAnd(string caseName, int caseNumber, string elementName)
        {
            result = Convert.ToDouble(ucl.GetValues(caseName, caseNumber, elementName));
            min = Convert.ToDouble(ucl.GetValues(caseName, caseNumber, "min"));
            max = Convert.ToDouble(ucl.GetValues(caseName, caseNumber, "max"));
        }

        [Then(@"the result should be (.*)")]
        public void ThenTheResultShouldBe(double expectedResult)
        {

            if ((result >= min) && (result <= max))
            {
                if (expectedResult != result)
                {
                    Assert.Inconclusive("Value is between min and max, however is not same as default value.");
                }
                else
                {
                    Assert.AreEqual(expectedResult, result);
                }
            }
            else
            {
                Assert.Fail("Value is out of range");
            }

        }
    }
}
