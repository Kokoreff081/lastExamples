using System;
using System.Collections.Generic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using IGIT.ASP.WebUI.DataAccess;
using IGIT.ASP.WebUI.Models;
using Kendo.Mvc;
using Kendo.Mvc.Extensions;
using Kendo.Mvc.UI;
using MoreLinq;

namespace IGIT.ASP.WebUI.Controllers
{
    public class ExplosionController : Controller
    {
        // GET: Explosion
        public ActionResult Index()
        {
            List<string> fss = new List<string>();
            List<string> fz = new List<string>();
            List<string> pue = new List<string>();
            List<RoomExpl> rooms = new List<RoomExpl>();
            using (AdmShipDataContext dc = new AdmShipDataContext())
            {
                var tFss = dc.GetFssCategories().ToList();
                foreach (var t in tFss) {
                    fss.Add(t.categoryFSS);
                }
                var tFz = dc.GetFzCategories().ToList();
                foreach (var t in tFz)
                {
                    fz.Add(t.categoryFZ);
                }
                var tPue = dc.GetPueCategories().ToList();
                foreach (var t in tPue)
                {
                    pue.Add(t.categoryPUE);
                }
                string query = @"select distinct r3.Name, r3.Назначение_отдельных_помещений+' ('+ r3.этаж+' этаж, пом.'+cast(cast(r3.номер_комнаты_по_плану as int) as nvarchar)+')' RoomName,
                                r3.Инвентарный_номер_здания BuildInvNumber from cyclone.dbo.ROOM_3D r3";
                var dbRooms = dc.ExecuteQuery<GetIdForExplosionResult>(query);
                foreach (var d in dbRooms) {
                    rooms.Add(new RoomExpl() { valField = d.Name, textField = d.BuildInvNumber+" ("+d.RoomName+")" });
                }
            }
            ViewData["fss"] = fss;//fssList;
            ViewData["fz"] = fz;//fzList;
            ViewData["pue"] = pue;//pueList;
            ViewData["rooms"] = rooms;
            return View();
        }

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Grid_Read([DataSourceRequest] DataSourceRequest request)
        {
            using (AdmShipDataContext dc = new AdmShipDataContext())
            {
                var res = dc.GetIdForExplosion()
                    .Select(s=>new ExplosionView() { Id = s.RoomId, Fss = s.CategoryFSS, Fz = s.CategoryFZ, Pue = s.CategoryPUE, RoomName = s.RoomName, BuildInvNumber= s.BuildInvNumber })
                    .OrderBy(o => o.RoomName)
                    .ToDataSourceResult(request);
                return Json(res);
            }
        }

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Grid_Update([DataSourceRequest] DataSourceRequest request, ExplosionView eV)
        {
            bool errorFlag = false;
            if (ModelState.IsValid)
            {

                using (AdmShipDataContext dc = new AdmShipDataContext())
                {
                    if (eV == null)
                    {
                        ModelState.AddModelError("EmptyField", "Не введены данные для сохранения");
                        errorFlag = true;
                        return GetUpDT(request, errorFlag, eV);
                    }
                    var tmp = dc.FireExplosions.Where(w => w.Name_Room == eV.Name).ToList();
                    if (tmp.Count > 0)
                    {
                        ModelState.AddModelError("AlreadyExist", "Такая связь уже существует!");
                        errorFlag = true;
                        return GetUpDT(request, errorFlag, eV);
                    }

                    dc.ExplosionInsertUpdateDelete(eV.Id, eV.Fss, eV.Fz, eV.Pue, eV.Name, "Update");
                }
            }

            return GetUpDT(request, errorFlag, eV);
        }

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Grid_Create([DataSourceRequest] DataSourceRequest request, ExplosionView eV)
        {
            bool errorFlag = false;
            if (ModelState.IsValid)
            {
                try
                {
                    using (AdmShipDataContext dc = new AdmShipDataContext())
                    {
                        if (eV == null)
                        {
                            ModelState.AddModelError("EmptyField", "Не введены данные для сохранения!");
                            errorFlag = true;
                            return GetUpDT(request, errorFlag, eV);
                        }
                        var tmp = dc.FireExplosions.Where(w => w.OID == eV.Id).ToList();
                        if (tmp.Count > 0)
                        {
                            ModelState.AddModelError("AlreadyExist", "Такая связь уже существует!");
                            errorFlag = true;
                            return GetUpDT(request, errorFlag, eV);
                        }

                        eV.Id = dc.ExplosionInsertUpdateDelete(eV.Id, eV.Fss, eV.Fz, eV.Pue, eV.Name, "Insert");
                    }
                }
                catch (Exception ex) {
                    errorFlag = true;
                    ModelState.AddModelError("SqlError", "Невозможно добавить такую запись. Возможные причины: \n 1. Такая запись уже существует в базе данных.\n 2. Предложенная запись не соответствует требованиям базы данных.");
                    return GetUpDT(request, errorFlag, eV);
                }
            }
            return GetUpDT(request, errorFlag, eV);
        }

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Grid_Destroy([DataSourceRequest] DataSourceRequest request, ExplosionView eV)
        {
            bool errorFlag = false;
            if (ModelState.IsValid)
            {
                using (AdmShipDataContext dc = new AdmShipDataContext())
                {
                    dc.ExplosionInsertUpdateDelete(eV.Id, eV.Fss, eV.Fz, eV.Pue, eV.Name, "Delete");
                }
            }
            return GetUpDT(request, errorFlag, eV);
        }

        private JsonResult GetUpDT(DataSourceRequest request, bool errorFlag, ExplosionView eV)
        {
            using (AdmShipDataContext dc = new AdmShipDataContext())
            {
                if (eV != null)
                {
                    var ls = dc.GetExplosionById(eV.Id).ToList();

                    if (ls.Any())
                    {
                        var b = ls.First();
                        eV.Id = b.RoomId;
                        eV.Fss = b.CategoryFSS;
                        eV.Fz = b.CategoryFZ;
                        eV.Pue = b.CategoryPUE;
                        eV.BuildInvNumber = b.BuildInvNumber;
                        eV.RoomName = b.RoomName;
                    }
                }

                return Json(new[] { eV }.ToDataSourceResult(request, ModelState));
            }
        }
    }
}