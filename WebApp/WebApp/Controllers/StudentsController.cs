using ClosedXML.Excel;
using LinqToExcel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Validation;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;
using System.Web.UI.WebControls;
using WebApp.Models;

namespace WebApp.Controllers
{
    public class StudentsController : Controller
    {
        private Database1Entities db = new Database1Entities();

        // GET: Students
        public ActionResult Index()
        {
            var students = db.Students.Include(s => s.Class);
            return View(students.ToList());
        }
        [HttpPost]
        public ActionResult ImportExcel(HttpPostedFileBase postedFile)
        {
            if (postedFile != null)
            {
                try
                {
                    string fileExtension = Path.GetExtension(postedFile.FileName);
                    ViewBag.Message = "fileExtension" + fileExtension;
                    //Validate uploaded file and return error.
                    if (fileExtension != ".xls" && fileExtension != ".xlsx")
                    {
                        ViewBag.Message = "Please select the excel file with .xls or .xlsx extension";
                        return View();
                    }

                    string folderPath = Server.MapPath("~/UploadedFiles/");
                    //Check Directory exists else create one
                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }

                    //Save file to folder
                    var filePath = folderPath + Path.GetFileName(postedFile.FileName);
                    postedFile.SaveAs(filePath);
                   
                   
                                       //Get file extension

                                        string excelConString = "";

                                        //Get connection string using extension 
                                        switch (fileExtension)
                                        {
                                            //If uploaded file is Excel 1997-2007.
                                            case ".xls":
                                                excelConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES'";
                                                break;
                                            //If uploaded file is Excel 2007 and above
                                            case ".xlsx":
                                                excelConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES'";
                                                break;
                                        }
                                        //Read data from first sheet of excel into datatable
                                        DataTable dt = new DataTable();
                                        excelConString = string.Format(excelConString, filePath);
                                    ViewBag.Message = "excelConString" + excelConString;
                                    using (OleDbConnection excelOledbConnection = new OleDbConnection(excelConString))
                                            {
                                            using (OleDbCommand excelDbCommand = new OleDbCommand())
                                            {
                                                using (OleDbDataAdapter excelDataAdapter = new OleDbDataAdapter())
                                                {
                                                    excelDbCommand.Connection = excelOledbConnection;

                                                    excelOledbConnection.Open();
                                                    //Get schema from excel sheet
                                                    DataTable excelSchema = GetSchemaFromExcel(excelOledbConnection);
                                                    //Get sheet name
                                                    string sheetName = excelSchema.Rows[0]["TABLE_NAME"].ToString();
                                                    excelOledbConnection.Close();

                                                    //Read Data from First Sheet.
                                                    excelOledbConnection.Open();
                                                    excelDbCommand.CommandText = "SELECT * From [" + sheetName + "]";
                                                    excelDataAdapter.SelectCommand = excelDbCommand;
                                                    //Fill datatable from adapter
                                                    excelDataAdapter.Fill(dt);
                                                    excelOledbConnection.Close();
                                                }
                                            }
                                        }

                                        //Insert records to Employee table.
                                        using (var context = new Database1Entities())
                                        {
                                            //Loop through datatable and add employee data to employee table. 
                                            foreach (DataRow row in dt.Rows)
                                            {
                                                
                                                context.Students.Add(GetStudentFromExcelRow(row));
                                            }

                                            context.SaveChanges();
                                        }
                                        ViewBag.Message = "Data Imported Successfully.";
                                    }
                                    catch (Exception ex)
                                    {
                                        ViewBag.Message = "Catch"+ ex.Message;
                                    }
                                }
                                else
                                {
                                    ViewBag.Message = "Please select the file first to upload.";
                                }
            return View();
        }

        private static DataTable GetSchemaFromExcel(OleDbConnection excelOledbConnection)
        {
            return excelOledbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
        }

        //Convert each datarow into employee object
        private Student GetStudentFromExcelRow(DataRow row)
        {
            return new Student
            {
                StID = row[0].ToString(),
                StName = row[1].ToString(),
                Password = row[2].ToString(),
                PortalID = row[3].ToString(),
                ClassID = row[4].ToString(),
                Email = row[5].ToString(),
                Phone = row[6].ToString(),
            };
        }

        public ActionResult ExportToExcel()
        {
           
            DataTable dt = new DataTable("Students");
            dt.Columns.AddRange(new DataColumn[7] { new DataColumn("StId"),
                                            new DataColumn("StName"),
                                            new DataColumn("Password"),
                                            new DataColumn("PortalId"),
                                            new DataColumn("ClassName"),
                                            new DataColumn("Email"),
                                            new DataColumn("Phone")
            });

            var student = from Students in db.Students.Take(10)
                            select Students;

            foreach (var st in student)
            {
                dt.Rows.Add(st.StID,st.StName,st.Password,st.PortalID,st.Class.ClassName,st.Email,st.Phone);
            }

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);
                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Grid.xlsx");
                }
            }
            return View("Index");
        }

        // GET: Students/Details/5
        public ActionResult Details(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Student student = db.Students.Find(id);
            if (student == null)
            {
                return HttpNotFound();
            }
            return View(student);
        }

        // GET: Students/Create
        public ActionResult Create()
        {
            ViewBag.ClassID = new SelectList(db.Classes, "ClassID", "ClassName");
            return View();
        }

        // POST: Students/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "StID,StName,Password,PortalID,ClassID,Email,Phone")] Student student)
        {
            if (ModelState.IsValid)
            {
                db.Students.Add(student);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.ClassID = new SelectList(db.Classes, "ClassID", "ClassName", student.ClassID);
            return View(student);
        }

        // GET: Students/Edit/5
        public ActionResult Edit(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Student student = db.Students.Find(id);
            if (student == null)
            {
                return HttpNotFound();
            }
            ViewBag.ClassID = new SelectList(db.Classes, "ClassID", "ClassName", student.ClassID);
            return View(student);
        }

        // POST: Students/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "StID,StName,Password,PortalID,ClassID,Email,Phone")] Student student)
        {
            if (ModelState.IsValid)
            {
                db.Entry(student).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.ClassID = new SelectList(db.Classes, "ClassID", "ClassName", student.ClassID);
            return View(student);
        }

        // GET: Students/Delete/5
        public ActionResult Delete(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Student student = db.Students.Find(id);
            if (student == null)
            {
                return HttpNotFound();
            }
            return View(student);
        }

        // POST: Students/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string id)
        {
            Student student = db.Students.Find(id);
            db.Students.Remove(student);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
