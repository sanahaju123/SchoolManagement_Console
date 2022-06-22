using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using ClosedXML.Excel;
using Grpc.Core;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System.Web.Mvc;
using System.Reflection;

namespace School_Management
{
    public class Program
    {
        private static List<School> school = new List<School>();
        public static void Main()
        {
            FileInfo fileInfo = new FileInfo(AppDomain.CurrentDomain.BaseDirectory);
            string parentDir = fileInfo.Directory.Parent.Parent.Parent.Parent.ToString();
            string path = Path.Combine(parentDir, @"SchoolManagement/Assets/ResultSheet.xlsx");
            string pathTxt = Path.Combine(parentDir, @"SchoolManagement/Assets/Test.txt");


            ClearText(pathTxt, FileMode.Truncate);

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("School Details");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "Id";
                worksheet.Cell(currentRow, 2).Value = "SchoolName";
                worksheet.Cell(currentRow, 3).Value = "Address";
                worksheet.Cell(currentRow, 4).Value = "Count of Students";
                worksheet.Cell(currentRow, 5).Value = "Date of Inauguration";
                for (int i = 2; i < 7; i++)
                {
                    Console.Write("Id: ");
                    int ID = int.Parse(Console.ReadLine());

                    Console.Write("SchoolName: ");
                    string name = Console.ReadLine();

                    Console.Write("Address: ");
                    string address = Console.ReadLine();

                    Console.Write("NumberOfstudents: ");
                    long count = int.Parse(Console.ReadLine());

                    Console.Write("DateofInauguration: ");
                    DateTime date = DateTime.Parse(Console.ReadLine());

                    School _schoolDetails = new School(ID, name, address, count, date);
                    List<School> _school = new List<School>();

                    AddSchoolDetails(worksheet, _schoolDetails, _school, i, workbook, path);

                }
                SerializeData(pathTxt, school);
                DeserializeData(pathTxt);

            }
        }

        /// <summary>
        /// Add at least 5 School details using List generic collection. <userinput>
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="_schoolDetails"></param>
        /// <param name="_school"></param>
        /// <param name="i"></param>
        public static bool AddSchoolDetails(IXLWorksheet worksheet, School _schoolDetails, List<School> _school, int i, XLWorkbook workbook, string path)
        {
            bool res = false;
            try
            {
                _school.Add(_schoolDetails);

                worksheet.Cell(i, 1).Value = _school[0].SchoolId;
                worksheet.Cell(i, 2).Value = _school[0].SchoolName;
                worksheet.Cell(i, 3).Value = _school[0].Address;
                worksheet.Cell(i, 4).Value = _school[0].NumberOfstudents;
                worksheet.Cell(i, 5).Value = _school[0].DateofInauguration;
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(path);
                    var content = stream.ToArray();
                }
                school.Add(_schoolDetails);
                res = true;
            }
            catch (Exception)
            {
                return res; ;
            }
            return res;
        }

        /// <summary>
        /// Save all School details in excel file and Serialize School List object in JSON format. 
        /// </summary>
        /// <param name="path"></param>
        /// <param name="pathTxt"></param>
        /// <param name="workbook"></param>
        /// <param name="_school"></param>
        public static bool SerializeData(string pathTxt, List<School> school)
        {
            bool res = false;
            try
            {
                var jsonValue = JsonConvert.SerializeObject(school);
                SaveTextFile(pathTxt, jsonValue);
                res = true;
            }
            catch (Exception)
            {
                return res; ;
            }
            return res;
        }

        /// <summary>
        /// JSON format save in text file. 
        /// </summary>
        /// <param name="pathTxt"></param>
        /// <param name="jsonValue"></param>
        public static bool SaveTextFile(string pathTxt, string jsonValue)
        {
            bool res = false;
            try
            {
                string text = File.ReadAllText(pathTxt);
                using (StreamWriter sw = File.AppendText(pathTxt))
                {
                    sw.WriteLine(jsonValue);
                }
                res = true;
            }
            catch (Exception)
            {
                return res; ;
            }
            return res;
        }

        /// <summary>
        /// Deserialize the fetched School list object. 
        /// </summary>
        /// <param name="pathTxt"></param>
        public static bool DeserializeData(string pathTxt)
        {
            bool res = false;
            try
            {
                string txt = File.ReadAllText(pathTxt);
                var values = JsonConvert.DeserializeObject<List<School>>(txt);
                DisplayAllDetails(values);
                res = true;
            }
            catch (Exception)
            {
                return res; ;
            }
            return res;
        }



        /// <summary>
        /// Show details of School in descending order of name
        /// </summary>
        /// <param name="values"></param>
        public static bool DisplayAllDetails(List<School> values)
        {
            bool res = false;
            try
            {
                values.Reverse();
                foreach (School skl in values)
                {
                    Console.WriteLine(skl.SchoolName);
                }
                res = true;
            }
            catch (Exception)
            {
                return res; ;
            }
            return res;
        }


        /// <summary>
        /// Empty text file.
        /// </summary>
        /// <param name="pathText"></param>
        /// <param name="fileMode"></param>
        public static bool ClearText(string pathText, FileMode fileMode)
        {
            bool res = false;
            try
            {

                using (var str = new FileStream(pathText, fileMode))
                {
                    using (var writer = new StreamWriter(str))
                    {
                        writer.Write("");
                    }
                }
                res = true;
            }
            catch (Exception ex)
            {
                return res;
            }
            return res;
        }
    }
    }



