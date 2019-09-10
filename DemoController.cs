using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ExcelFileDownload.Models;
using System.Data;
using Spire.Xls;

namespace ExcelFileDownload.Controllers
{
    public class DemoController : Controller
    {
        // GET: Demo
        public ActionResult Index()
        {
            return View();
        }
        //for fetching the book details 
        public PartialViewResult BookDetails()
        {
            var bookDetails=new List<BookDetail>();
            using (MKDBEntities DBContext = new MKDBEntities())
            {
                bookDetails = DBContext.BookDetails.ToList();
            }
            return PartialView(bookDetails);
        }
        public PartialViewResult CourseDetails()
        {
            var courseDetails = new List<Course>();
            using (MKDBEntities DBContext = new MKDBEntities())
            {
                courseDetails = DBContext.Courses.ToList();

            }
            return PartialView(courseDetails);

        }

        public PartialViewResult TeacherDetails()
        {
            var teacherDetails = new List<Teacher>();
      
            using (MKDBEntities DBContext = new MKDBEntities())
            {
                teacherDetails = DBContext.Teachers.ToList();
            }
            return PartialView(teacherDetails);
        }
     
        public ActionResult DownloadExcel()
        {

            BusinessLayer BAL = new BusinessLayer();
            var workbook = new Spire.Xls.Workbook();

            workbook.LoadFromFile(Server.MapPath("~/DetailFormatInExcel/DetailsFormat.xlsx"));
            var worksheet1 = workbook.Worksheets[0];
            var worksheet2 = workbook.Worksheets[1];
            byte[] array = null;
            var dt1 = BAL.GetXlsTableBooks();
            worksheet1.InsertDataTable(dt1, false, 3, 1);
            var dt2 = BAL.GetXlsTableCourse();
            worksheet2.InsertDataTable(dt2, false, 3, 1);
            var dt3 = BAL.GetXlsTableTeacher();
            worksheet2.InsertDataTable(dt3, false, 3, 3);
          
            using (var ms = new System.IO.MemoryStream())
            {
                workbook.SaveToStream(ms, FileFormat.Version2010);
                ms.Seek(0, System.IO.SeekOrigin.Begin);
                array = ms.ToArray();
            }

            return File(array, "application / vnd.openxmlformats - officedocument.spreadsheetml.sheet"," Detail.xlsx");
        }

    }

}