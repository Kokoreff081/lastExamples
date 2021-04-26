using System;
using System.Collections.Generic;
using System.Data.Common;
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
    public class SubdivBuildEditController : Controller
    {
        private List<Build> builds = AbstractDataAccess.SelectA(@"select distinct substr(v_gis_os1.inv_number, 1, 6) AS inv_number, name from v_gis_os1 where code_subsection = 210 order by name", r => Build.Get(r));
        private List<Subdiv> subdivs = AbstractDataAccess.SelectA(@"select distinct department, dep_full_name from v_gis_pers order by dep_full_name", r => Subdiv.Get(r));
        // GET: SubdivBuildEdit
        public ActionResult Index()
        {
            ViewData["builds"] = builds;
            ViewData["subdivs"] = subdivs;
            ViewData["defBuild"] = builds.First();
            ViewData["defSubdiv"] = subdivs.First();
            return View("SubBuildEdit");
        }

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Grid_Read([DataSourceRequest] DataSourceRequest request)
        {
            using (AdmShipDataContext dc = new AdmShipDataContext())
            {
                var result = dc.GetBuildSubdivConnection().Select(r => new SubBuildEditModel()
                {
                    buildId = r.id,
                    buildNumber = r.buildNumber,
                    buildName = r.buildName,
                    subdivisionName = r.subdivisionName,
                    subdivId = r.subdivId
                }).ToList();

                return Json(result.OrderBy(o => o.nameOfBuild).ThenBy(t => t.nameOfSubdiv).ToDataSourceResult(request));
            }
        }
        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Grid_Update([DataSourceRequest] DataSourceRequest request, SubBuildEditModel build)
        {
            return UpdateCreate(request, build, "Update");
        }

        private ActionResult UpdateCreate(DataSourceRequest request, SubBuildEditModel build, string type)
        {
            bool errorFlag = false;
            if (ModelState.IsValid)
            {
                using (AdmShipDataContext dc = new AdmShipDataContext())
                {
                    if (build == null)
                    {
                        ModelState.AddModelError("EmptyField", "Не заполнено одно или несколько из обязательных полей!");
                        errorFlag = true;
                        return GetUpDT(request, errorFlag, build);
                    }

                    if (String.IsNullOrEmpty(build?.InvNumber))
                    {
                        ModelState.AddModelError("EmptyBuild", "Не выбрано здание!");
                        errorFlag = true;
                        return GetUpDT(request, errorFlag, build);
                    }

                    if (build.SubDiv == 0)
                    {
                        ModelState.AddModelError("EmptySubdiv", "Не выбрано подразделение!");
                        errorFlag = true;
                        return GetUpDT(request, errorFlag, build);
                    }

                    if (!subdivs.Any(r => r.Id == build.SubDiv))
                    {
                        errorFlag = true;
                        ModelState.AddModelError("NotExistSubdiv", "Такого подразделения не существует!");
                        return GetUpDT(request, errorFlag, build);
                    }

                    if (!builds.Any(r => r.InvNumber == build.InvNumber))
                    {
                        errorFlag = true;
                        ModelState.AddModelError("NotExistBuild", "Такого здания не существует!");
                        return GetUpDT(request, errorFlag, build);
                    }

                    var ls = dc.GetBuildSubDiv(build.InvNumber, build.SubDiv);

                    if (ls.Any())
                    {
                        errorFlag = true;
                        ModelState.AddModelError("Exists", "Такая связь существует!");
                        return GetUpDT(request, errorFlag, build);
                    }


                    var b = builds.First(r => r.InvNumber == build.InvNumber);

                    try
                    {
                        dc.SubdivBuildEdit(build.buildId, build.SubDiv, b.Name, b.InvNumber, type);
                    }
                    catch (Exception e)
                    {
                        errorFlag = true;
                        ModelState.AddModelError("Error", e.Message);
                    }
                }
            }

            return GetUpDT(request, errorFlag, build);
        }

        private JsonResult GetUpDT(DataSourceRequest request, bool errorFlag, SubBuildEditModel build)
        {
            using (AdmShipDataContext dc = new AdmShipDataContext())
            {
                if (build != null)
                {
                    var ls = dc.GetBuildSubDiv(build.InvNumber, build.SubDiv);

                    if (ls.Any())
                    {
                        var b = ls.First();
                        build.buildId = b.id;
                        build.buildNumber = b.buildNumber;
                        build.buildName = b.buildName;
                        build.subdivisionName = b.subdivisionName;
                        build.subdivId = b.subdivId;
                    }
                }

                return Json(new[] { build }.ToDataSourceResult(request, ModelState));
            }
        }

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Grid_Create([DataSourceRequest] DataSourceRequest request, SubBuildEditModel build)
        {
            return UpdateCreate(request, build, "Insert");
        }
        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Grid_Destroy([DataSourceRequest] DataSourceRequest request, SubBuildEditModel build)
        {
            bool errorFlag = false;
            DbTransaction dbt = null;
            dynamic result = null;
            if (ModelState.IsValid)
            {
                try
                {
                    int subdivId = 0;
                    int tmpS = 0;
                    if (Int32.TryParse(build.nameOfSubdiv, out tmpS))
                        subdivId = Int32.Parse(build.nameOfSubdiv);
                    string bnumber = "";
                    foreach (var b in builds)
                    {
                        if (b.buildName == build.nameOfBuild)
                            bnumber = b.InvNumber;
                    }
                    using (AdmShipDataContext dc = new AdmShipDataContext())
                    {
                        if (dc.Connection.State == System.Data.ConnectionState.Closed)
                        {
                            dc.Connection.Open();
                        }
                        dbt = dc.Connection.BeginTransaction(System.Data.IsolationLevel.Serializable);
                        dc.Transaction = dbt;
                        dc.SubdivBuildEdit(build.buildId, subdivId, build.nameOfBuild, bnumber, "Delete");
                        dbt.Commit();

                    }
                }
                catch (Exception ex)
                {
                    errorFlag = true;
                    ModelState.AddModelError("Del", "Невозможно удалить запись!");
                }
            }
            return GetUpDT(request, errorFlag, build);

        }

    }
}