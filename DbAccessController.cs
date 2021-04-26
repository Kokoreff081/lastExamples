using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Web;
using System.Web.Configuration;
using System.Web.Mvc;
using IGIT.ASP.WebUI.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using static System.Web.Helpers.Json;
using System.Data.Linq;
using System.Globalization;
using System.Net;
using Microsoft.EntityFrameworkCore;
using IGIT.ASP.WebUI.DataAccess;




namespace IGIT.ASP.WebUI.Controllers
{
    public class DbAccessController : Controller
    {
        //OracleDbContext db2 = new OracleDbContext();
        MsAdmShip db3 = new MsAdmShip(@""+WebConfigurationManager.ConnectionStrings["IGIT.ASP.WebUI.Properties.Settings.AdmShipConnectionString"] +"");
        //AdmShipDataContext db = new AdmShipDataContext();
        //CycloneDataContext dc = new CycloneDataContext();
        [HttpPost]
        public JsonResult GetEmployees()
        {
            Dictionary<string, List<GisRasp>> toJson = new Dictionary<string, List<GisRasp>>();
            //string json = "";
            var employees = AbstractDataAccess.SelectA(@"select * from V_GIS_PERS order by FNAME", r => new GisRasp(r));//where PROF_NAME not likec '%Начальник%'db2.Users.Where(c => !c.profName.Contains("Начальник")).OrderBy(o=>o.fName).ToList();
            //toJson.Add("list", employees);
            //json = JsonConvert.SerializeObject(toJson, Formatting.Indented);
            return Json(employees);
        }
        
        [HttpPost]
        public JsonResult GetChiefs()
        {
            Dictionary<string, List<GisRasp>> toJson = new Dictionary<string, List<GisRasp>>();
            string json = "";
            var chiefs = AbstractDataAccess.SelectA(@"select * from V_GIS_PERS where PROF_NAME like 'Начальник%' or PROF_NAME like 'Главный%' or PROF_NAME like 'Генеральный%' or PROF_NAME like 'Заместитель%' order by FNAME", r => new GisRasp(r));//db2.Users.Where(c=> c.profName.Contains("Начальник")).OrderBy(o => o.fName).ToList();
            toJson.Add("list", chiefs);
            json = JsonConvert.SerializeObject(toJson, Formatting.Indented);
            return Json(json);
        }
        [HttpPost]
        public JsonResult GetBuildings()
        {
            using (CycloneDataContext dc = new CycloneDataContext())
            {
                var buildings = dc.GetBuildingList().ToList();
                return Json(buildings);
            }
            //var buildingsAll = AbstractDataAccess.SelectA(@"select distinct inv_number, name, kc_up from v_gis_os1 where code_subsection = 210 order by inv_number", r => GisOs1.Get4(r));
            //var buildings = buildingsC.Join(buildingsAll, bc => bc.InvNumber.Substring(0,6), ba => ba.invNumber.Substring(0,6), (bc, ba) => new { name = bc.ObjName, invNumber = bc.InvNumber, codeSubdiv = ba.codeSubdivision }).ToList();
            
        }
        [HttpPost]
        public JsonResult GetBuildings3D()
        {
            using (AdmShipDataContext db = new AdmShipDataContext())
            {
                var buildings = db.GetBuildings3d().ToList();
                return Json(buildings);
            }
            
        }
        [HttpPost]
        public JsonResult GetSubdivisions()
        {
            Dictionary<string, List<GisOs1>> toJson = new Dictionary<string, List<GisOs1>>();
            //string json = "";
            //var subdivisionsB = (from s in db3.GetTable<Room>()
            //                    select new { codeSubdivision = s.departId, subdivId = s.subdivisionId, subdivisionName = s.subdivisionName}).Distinct().ToList();
            var subdivisionsU = AbstractDataAccess.SelectA(@"select distinct department, dep_short_name, dep_full_name from v_gis_pers order by department", r => GisRasp.Get3(r));
            //db2.Users.Select(s => new { codeSubdivision = s.departmentId, subdivisionName = s.departmentFullName, idChief = s.Id }).Distinct().ToList();
            //var subdivisions = subdivisionsB
            //    .Join(subdivisionsU, b => b.codeSubdivision, u => u.departmentId, (b, u) => new
            //    {
            //        codeSubdivision = b.subdivId,
            //        subdivisionName = b.subdivisionName,
            //        depId = u.departmentId,
            //    })
            //    .OrderBy(o => o.subdivisionName)
            //    .GroupBy(g=>g.codeSubdivision)
            //    .ToList();
            //toJson.Add("list", subdivisions);
            //json = JsonConvert.SerializeObject(toJson, Formatting.Indented);
            return Json(subdivisionsU);
        }

        [HttpPost]
        public JsonResult GetSubdivisions2D() {
            using (AdmShipDataContext db = new AdmShipDataContext())
            {
                var subdivisionsU = db.GetSubdivisions2D().ToList();
                return Json(subdivisionsU);
            }
        }

        [HttpPost]
        public JsonResult GetRooms3D()
        {
            using (AdmShipDataContext db = new AdmShipDataContext())
            {
                var rooms = db.GetRooms().OrderBy(o => o.roomId).ToList();
            
                JsonSerializerSettings settings = new JsonSerializerSettings();
                settings.ReferenceLoopHandling = ReferenceLoopHandling.Ignore;
            
                JsonResult jsonResult = Json(rooms);
                jsonResult.MaxJsonLength = int.MaxValue;
                return jsonResult;
            }
        }
        [HttpPost]
        public JsonResult GetRooms2D()
        {
            using (AdmShipDataContext db = new AdmShipDataContext())
            {
               var rooms = db.GetRooms2D().OrderBy(o => o.roomId).ToList();
            
                JsonSerializerSettings settings = new JsonSerializerSettings();
                settings.ReferenceLoopHandling = ReferenceLoopHandling.Ignore;

                JsonResult jsonResult = Json(rooms);
                jsonResult.MaxJsonLength = int.MaxValue;
                return jsonResult;
            }
        }
        public JsonResult GetRoomEmployees(int roomId, int subdiv) {
            
            using (AdmShipDataContext db = new AdmShipDataContext())
            {
                var rooms = db.GetRoomsForEdit(roomId, subdiv).ToList();
                return Json(rooms);
            }
            
        }

        //методы для функционала экспорта объектов в dwg формат
        [HttpPost]
        public JsonResult GetGeoprocUrl(string val) {
            string url = url = WebConfigurationManager.AppSettings[val];
            return Json(url);
        }
        [HttpPost]
        public void GetDwgResult(string url, string filename) {
            WebClient wc = new WebClient();
            wc.DownloadFile(url, filename);
        }
        
    }
}