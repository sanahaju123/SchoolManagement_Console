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
            throw new NotImplementedException();
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
            throw new NotImplementedException();
        }
        
        /// <summary>
        /// JSON format save in text file. 
        /// </summary>
        /// <param name="pathTxt"></param>
        /// <param name="jsonValue"></param>
        public static bool SaveTextFile(string pathTxt, string jsonValue)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Deserialize the fetched School list object. 
        /// </summary>
        /// <param name="pathTxt"></param>
        public static bool DeserializeData(string pathTxt)
        {
            throw new NotImplementedException();
        }


        /// <summary>
        /// Show details of School in descending order of name
        /// </summary>
        /// <param name="values"></param>
        public static bool DisplayAllDetails(List<School> values)
        {
            throw new NotImplementedException();
        }


        /// <summary>
        /// Empty text file.
        /// </summary>
        /// <param name="pathText"></param>
        /// <param name="fileMode"></param>
        public static bool ClearText(string pathText, FileMode fileMode)
        {
            throw new NotImplementedException();
        }
    }
}



