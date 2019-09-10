using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ExcelFileDownload.Models;
using LinqToExcel;
using System.Data;
using System.Reflection;

namespace ExcelFileDownload.Controllers
{
    public class ExcelDemoController : Controller
    {
        // GET: ExcelDemo
        public ActionResult ExcelUpload()
        {
            return View();
        }
        [HttpPost]
        public ActionResult UploadExcel(EmployeeInfo objEmpDetail, HttpPostedFileBase FileUpload)
        {
            var dataTable = CreateDataTable();
            EmployeeDBEntities objEntity = new EmployeeDBEntities();
            string data = "";
            if (FileUpload != null)
            {
                if (FileUpload.ContentType == "application/vnd.ms-excel" || FileUpload.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                {
                    string filename = FileUpload.FileName;

                    if (filename.EndsWith(".xlsx"))
                    {
                        string targetpath = Server.MapPath("~/DetailFormatInExcel/");
                        FileUpload.SaveAs(targetpath + filename);
                        string pathToExcelFile = targetpath + filename;

                        string sheetName = "Sheet1";

                        var excelFile = new ExcelQueryFactory(pathToExcelFile);
                        var empDetails = from a in excelFile.Worksheet<EmployeeInfo>(sheetName) select a;
                        foreach (var a in empDetails)
                        {
                            if (a.EmployeeNo != null)
                            {

                                DateTime? myBirthdate = null;


                                if (a.MobileNo.Length > 12)
                                {
                                    data = "Phone number should be 10 to 12 disit";
                                    ViewBag.Message = data;

                                }

                                myBirthdate = Convert.ToDateTime(a.DateOfBirth);


                                DataTable resullt = PostExcelData(dataTable, a.EmployeeNo, a.FirstName, a.LastName, myBirthdate, a.Address, a.MobileNo, a.PostelCode, a.EmailId);
                                
                            }

                            else
                            {
                                data = a.EmployeeNo + "Some fields are null, Please check your excel sheet";
                                ViewBag.Message = data;
                                return View("ExcelUpload");
                            }
                           
                        }
                    }

                    else
                    {
                        data = "This file is not valid format";
                        ViewBag.Message = data;
                    }
                    return View("ExcelUpload");

                }
                else
                {

                    data = "Only Excel file format is allowed";

                    ViewBag.Message = data;
                    return View("ExcelUpload");

                }

            }
            else
            {

                if (FileUpload == null)
                {
                    data = "Please choose Excel file";
                }

                ViewBag.Message = data;
                return View("ExcelUpload");

            }
        }
        
        public DataTable PostExcelData(DataTable dataTable, int employeeNo,string firstName, string lastName, DateTime? dateOfBirth, string address,string mobileNo,string postelCode, string emailId)
        {
            dataTable= InsertDatainDataTable(dataTable, employeeNo, firstName, lastName, dateOfBirth, address, mobileNo, postelCode, emailId);
            //EmployeeDBEntities DbEntity = new EmployeeDBEntities();
            //var InsertExcelData = DbEntity.usp_InsertNewEmployeeDetails(employeeNo,firstName, lastName, dateOfBirth,address,mobileNo,postelCode, emailId);
            return dataTable;
        }

        public DataTable CreateDataTable()
        {
            Type type = typeof(EmployeeInfo);
            var properties = type.GetProperties();

            DataTable dataTable = new DataTable();
            foreach (PropertyInfo info in properties)
            {
                dataTable.Columns.Add(new DataColumn(info.Name, Nullable.GetUnderlyingType(info.PropertyType) ?? info.PropertyType));
            }
            return dataTable;
            
        }

        public DataTable InsertDatainDataTable(DataTable dataTable, int employeeNo, string firstName, string lastName, DateTime? dateOfBirth, string address, string mobileNo, string postelCode, string emailId)
        {
            Type type = typeof(EmployeeInfo);
            var properties = type.GetProperties();
            dataTable.Rows.Add(employeeNo, firstName, lastName, dateOfBirth, address, mobileNo, postelCode, emailId);
            return dataTable;
        }
    }
}