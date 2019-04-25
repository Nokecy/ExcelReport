using System.Collections.Generic;

namespace EPPlus多行重复渲染
{
    public static class StudentLogic
    {
        public static List<StudentInfo> GetList()
        {
            List<StudentInfo> list = new List<StudentInfo>();

            list.Add(new StudentInfo() { Class = "一班", Name = "ab", Gender = true, RecordNo = "ab", Phone = "1", Email = "cn" });
            list.Add(new StudentInfo() { Class = "二班", Name = "ab", Gender = false, RecordNo = "ab", Phone = "1", Email = "cn" });
            list.Add(new StudentInfo() { Class = "一班", Name = "ab", Gender = true, RecordNo = "ab", Phone = "15", Email = "cn" });
            list.Add(new StudentInfo() { Class = "一班", Name = "ab", Gender = true, RecordNo = "ab", Phone = "15", Email = "cn" });

            return list;
        }
    }
}