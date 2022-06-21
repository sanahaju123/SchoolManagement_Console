using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Text;

namespace School_Management
{
    public class School : IComparable<School>
    {
        public int SchoolId { get; set; }
        public String SchoolName { get; set; }
        public String Address { get; set; }
        public long NumberOfstudents { get; set; }
        public DateTime DateofInauguration { get; set; }
        public School(int schoolId, string name, string address, long count, DateTime date)
        {
            this.SchoolId = schoolId;
            this.SchoolName = name;
            this.Address = address;
            this.NumberOfstudents = count;
            this.DateofInauguration = date;
        }

        

        public School()
        {
        }

        public int CompareTo([AllowNull] School other)
        {
            return this.SchoolName.CompareTo(other.SchoolName);
        }
    }
}
