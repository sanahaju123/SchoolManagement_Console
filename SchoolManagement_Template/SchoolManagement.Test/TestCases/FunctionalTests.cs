using ClosedXML.Excel;
using School_Management;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;

namespace SchoolManagement.Tests.TestCases
{
    public class FunctionalTests
    {
        // <summary>
        /// Creating referance variable of ITestOutputHelper and injecting in constructor
        /// </summary>
        private readonly ITestOutputHelper _output;
        private School _school;
        string path;
        string pathTxt;

        private static string type = "Functional";
        public FunctionalTests(ITestOutputHelper output)
        {
            FileInfo fileInfo = new FileInfo(AppDomain.CurrentDomain.BaseDirectory);
            string parentDir = fileInfo.Directory.Parent.Parent.Parent.Parent.ToString();
            path = Path.Combine(parentDir, @"SchoolManagement/Assets/ResultSheet.xlsx");
            pathTxt = Path.Combine(parentDir, @"SchoolManagement/Assets/Test.txt");

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
        /// Test for Serialization.
        /// </summary>
        /// <returns></returns>
        [Fact]
        public async Task<bool> TestFor_Serialization()
        {
            //Arrange
            bool res = false;
            string testName; string status;
            List<School> school = new List<School>();
            school.Add(_school);
            testName = CallAPI.GetCurrentMethodName();
            try
            {
                bool result = Program.SerializeData(pathTxt, school);
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


        /// <summary>
        /// Test for deserialization.
        /// </summary>
        /// <returns></returns>
        [Fact]
        public async Task<bool> TestFor_Deserialization()
        {
            //Arrange
            bool res = false;
            string testName; string status;
            testName = CallAPI.GetCurrentMethodName();
            try
            {
                bool result = Program.DeserializeData(pathTxt);

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


        /// <summary>
        /// test for save text file.
        /// </summary>
        /// <returns></returns>
        [Fact]
        public async Task<bool> TestFor_Save_TextFile()
        {
            //Arrange
            bool res = false;
            string testName; string status;
            string jsonValue = "";
            testName = CallAPI.GetCurrentMethodName();
            try
            {
                bool result = Program.SaveTextFile(pathTxt, jsonValue);
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

        /// <summary>
        /// Display All Details.
        /// </summary>
        /// <returns></returns>
        [Fact]
        public async Task<bool> TestFor_DisplayAllDetails()
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

        /// <summary>
        /// test for clear text file data.
        /// </summary>
        /// <returns></returns>
        [Fact]
        public async Task<bool> TestFor_ClearText()
        {
            //Arrange
            bool res = false;
            string testName; string status;
            FileMode fileMode = FileMode.Truncate;
            testName = CallAPI.GetCurrentMethodName();
            try
            {
                bool result = Program.ClearText(pathTxt, fileMode);
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

        /// <summary>
        /// test for Add school data.
        /// </summary>
        /// <returns></returns>
        [Fact]
        public async Task<bool> TestFor_AddSchoolDetails()
        {
            //Arrange
            bool res = false;
            string testName; string status;
            testName = CallAPI.GetCurrentMethodName();
            IXLWorksheet worksheet = null;
            XLWorkbook workbook = null;
            int i = 0;
            List<School> school = new List<School>();
            school.Add(_school);
            try
            {
                bool result = Program.AddSchoolDetails(worksheet, _school, school, i, workbook, path);
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

            //Assert
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

