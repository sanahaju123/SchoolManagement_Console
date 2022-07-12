using School_Management;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;

namespace SchoolManagement.Tests.TestCases
{
    public class ExceptionalTests
    {
        // <summary>
        /// Creating referance variable of ITestOutputHelper and injecting in constructor
        /// </summary>
        private readonly ITestOutputHelper _output;
        private School _school;
        private static string type = "Exception";
        public ExceptionalTests(ITestOutputHelper output)
        {
            _output = output;
            _school = new School()
            {
                SchoolId = 1,
                SchoolName = "MySchool",
                Address = "Address1",
                NumberOfstudents = 15,
                DateofInauguration = DateTime.Now
            };
        }

        /// <summary>
        /// If passing null will return false.
        /// </summary>
        /// <returns></returns>
        [Fact]
        public async Task<bool> TestFor_ReturnFalse_DisplayAllDetails()
        {
            //Arrange
            bool res = false;
            string testName; string status;
            List<School> school = new List<School>();
            school.Add(null);
            testName = CallAPI.GetCurrentMethodName();
            try
            {
                bool result = Program.DisplayAllDetails(school);
                //Act
                if (result == false)
                {
                    res = true;
                }
            }
            catch (Exception)
            {
                status = Convert.ToString(res);
                _output.WriteLine(testName + ":Failed");
                await CallAPI.saveTestResult(testName, status, type);
                return false;
            }
            ///Assert
            status = Convert.ToString(res);
            if (res == true)
            {
                _output.WriteLine(testName + ":Passed");
            }
            else
            {
                _output.WriteLine(testName + ":Failed");
            }
            await CallAPI.saveTestResult(testName, status, type);
            return res;
        }

    }
}
