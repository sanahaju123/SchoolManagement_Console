using School_Management;
using SchoolManagement;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;

namespace SchoolManagement.Tests.TestCases
{
    public class BoundaryTests
    {
        // <summary>
        /// Creating referance variable of ITestOutputHelper and injecting in constructor
        /// </summary>
        private readonly ITestOutputHelper _output;
        private School _school;
        private static string type = "Boundary";
        public BoundaryTests(ITestOutputHelper output)
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
        /// Validate Display All Details.
        /// </summary>
        /// <returns></returns>
        [Fact]
        public async Task<bool> TestFor_ValidateDisplayAllDetails()
        {
            //Arrange
            bool res = false;
            string testName; string status;
            List<School> school = new List<School>();
            school.Add(_school);
            testName = CallAPI.GetCurrentMethodName();
            try
            {
                bool result = Program.DisplayAllDetails(school);
                //Act
                if (result == true)
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
