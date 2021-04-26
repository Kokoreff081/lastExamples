using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Web;
using System.Web.Mvc;
using IGIT.ASP.WebUI.DataAccess;
using IGIT.ASP.WebUI.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System.Globalization;

namespace IGIT.ASP.WebUI.Controllers
{
    public static class PredicateBuilder
    {
        public static Expression<Func<T, bool>> True<T>() { return f => true; }
        public static Expression<Func<T, bool>> False<T>() { return f => false; }

        public static Expression<Func<T, bool>> Or<T>(this Expression<Func<T, bool>> expr1,
                                                            Expression<Func<T, bool>> expr2)
        {
            var invokedExpr = Expression.Invoke(expr2, expr1.Parameters.Cast<Expression>());
            return Expression.Lambda<Func<T, bool>>
                  (Expression.OrElse(expr1.Body, invokedExpr), expr1.Parameters);
        }

        public static Expression<Func<T, bool>> And<T>(this Expression<Func<T, bool>> expr1,
                                                             Expression<Func<T, bool>> expr2)
        {
            var invokedExpr = Expression.Invoke(expr2, expr1.Parameters.Cast<Expression>());
            return Expression.Lambda<Func<T, bool>>
                  (Expression.AndAlso(expr1.Body, invokedExpr), expr1.Parameters);
        }
    }

    public class ServLocationController : Controller
    {
        // GET: ServLocation
        [HttpPost]
        public JsonResult Index(string id, int? subdiv=null, int? chief = null, int? roomId = null, string data = "", int? emplId = null)
        {
            //interfereDataBases();
            TableToFront2 table = new TableToFront2();
            List<TableToFront2> lst = new List<TableToFront2>();
            List<TableToFront2> lstChief = new List<TableToFront2>();
            //List<ObjectToReturn2D> dict = new List<ObjectToReturn2D>();
            List<IsAnyConnectionResult> check = new List<IsAnyConnectionResult>();
            using (AdmShipDataContext dc = new AdmShipDataContext()) {
                check = dc.IsAnyConnection().ToList();
            }
            ObjectToReturn objToReturn = new ObjectToReturn();
            if (check.Any(x=>x.buildNumber==id))
            {
                List<GetRoomByAllFields2DResult> rooms = new List<GetRoomByAllFields2DResult>();
                if (data != "")
                {
                    List<FromFilter> parsedData = (List<FromFilter>)JsonConvert.DeserializeObject(data, typeof(List<FromFilter>));
                    string declare = "declare @buildNum nvarchar(6) = '" + id + "';";
                    if (subdiv == null)
                    {
                        declare += "declare @subDivId int;";
                    }
                    else
                        declare += "declare @subDivId int = " + subdiv + ";";
                    if (roomId == null)
                    {
                        declare += "declare @roomId int;";
                    }
                    else
                    {
                        declare += "declare @roomId int= " + roomId + ";";
                    }
                    if (chief == null)
                    {
                        declare += "declare @chiefId int;";
                    }
                    else
                    {
                        declare += "declare @chiefId int= " + chief + ";";
                    }
                    if (emplId == null)
                    {
                        declare += "declare @empId int;";
                    }
                    else
                    {
                        declare += "declare @empId int= " + emplId + ";";
                    }
                    string action = @"SELECT distinct r2.invNumber, r2.buildName, r7.SubdivisionName, r7.DepartmentId, r5.roomFloor, r1.roomId, r1.roomFilename3ds, r1.uniqueRoomNumber, r8.OBJECTID,
	                [dbo].GetChiefFio(r1.roomId, @subDivId) chiefFio, [dbo].GetEmpFio(r1.roomId, @subDivId) empFio from [dbo].new_Buildings r2 
	                left join [dbo].[new_SubdivBuild] r9 on (r9.buildNumber = r2.invNumber)
	                left join [dbo].new_Subdivisions r7 on(r9.subdivId = r7.DepartmentId)
	                join [dbo].new_Rooms r1  on(r2.buildId = r1.buildId)
                    left join new_EmployeesInRoom r3 on (r3.roomId = r1.roomId)
	                left join [dbo].new_RoomDetails r5 on (r1.roomId = r5.roomId) 
	                left join cyclone.dbo.ROOM_2D r8 on (r1.roomFilename3ds = r8.Name)
	                where (@buildNum is null or r2.invNumber = @buildNum) and 
                        (@subDivId is null or r7.DepartmentId = @subDivId)
                     and (@roomId is null or r1.roomId = @roomId)
                     and (@chiefId is null or r3.chiefId = @chiefId) and (@empId is null or r3.empId = @empId) and ";
                    string oper = "";
                    string action2 = declare + action;
                    //var predicate = PredicateBuilder.True<Room>();
                    //predicate = predicate.And(r => r.buildInvNumber.Contains(id) && r.subdivisionId == subdiv);//&& r.idChief == chief
                    string action3 = CreateQuery(action2, parsedData);
                    string query = SubstringQuery(action3);
                    using (AdmShipDataContext db3 = new AdmShipDataContext())
                    {
                        var rooms_pars = db3.ExecuteQuery<GetRoomByAllFields2DResult>(query).OrderBy(o => o.roomId).ToList();//db3.GetTable<Room>().FromSql($"SELECT * FROM dbo.rooms({action})").ToList();
                        rooms = rooms_pars;
                    }
                }
                else
                {
                    using (AdmShipDataContext db3 = new AdmShipDataContext())
                    {
                        if (id != "")
                            rooms = db3.GetRoomByAllFields2D(id, subdiv, roomId, chief, emplId).OrderBy(o => o.roomId).ToList();//GetTable<Room>().Where(r => r.subdivisionId == subdiv).OrderBy(o => o.uniqueRoomNumber).ToList();&& r.idChief == chief - добавим позже
                    }
                }
                foreach (var room in rooms)
                {
                    TableToFront2 tbl2 = new TableToFront2();
                    tbl2.OBJECTID = room.OBJECTID;
                    tbl2.roomFloor = room.roomFloor;
                    tbl2.roomNumber = room.uniqueRoomNumber;
                    tbl2.roomId = (int)room.roomId;
                    if (room.empFio != null)
                    {
                        tbl2.employeeFio = room.empFio;//.Split(',');
                    }
                    else
                    {
                        tbl2.employeeFio = /*new string[] {*/ string.Empty;
                    }
                    if (room.chiefFio != null)
                        tbl2.chiefFio = room.chiefFio;
                    else
                        tbl2.chiefFio = string.Empty;
                    tbl2.subdivisionId = room.DepartmentId;
                    tbl2.subdivision = room.subdivisionName;
                    tbl2.depId = room.DepartmentId;
                    tbl2.buildInvNumber = room.invNumber;
                    tbl2.buildName = room.buildName;
                    tbl2.filename3ds = room.roomFilename3ds;
                    lst.Add(tbl2);
                }

                foreach (var t in lst)
                {
                    if (t.subdivisionId == null)
                        t.isSubdivNull = 0;
                    else
                        t.isSubdivNull = 1;
                    if (t.OBJECTID == null)
                        t.isVisible = 0;
                    else
                        t.isVisible = 1;
                }

                objToReturn.flag = true;
                objToReturn.list = lst;
            }
            else {
                List<GetRoomByAllFields_oldResult> rooms = new List<GetRoomByAllFields_oldResult>();
                if (data != "")
                {
                    List<FromFilter> parsedData = (List<FromFilter>)JsonConvert.DeserializeObject(data, typeof(List<FromFilter>));
                    string declare = "declare @buildNum nvarchar(6) = '" + id + "';";
                    if (roomId == null)
                    {
                        declare += "declare @roomId int;";
                    }
                    else
                    {
                        declare += "declare @roomId int= " + roomId + ";";
                    }
                    string action = @"select distinct r1.roomId, r1.roomFilename3ds, r1.uniqueRoomNumber, r2.invNumber, r2.buildName, r5.roomFloor,  r8.OBJECTID, r7.SubdivisionName, r7.DepartmentId, [dbo].GetChiefFio(r1.roomId, r7.DepartmentId) chiefFio, [dbo].GetEmpFio(r1.roomId, r7.DepartmentId) empFio
	                                    from [dbo].new_Rooms r1
				                            join [dbo].new_Buildings r2  on(r2.buildId = r1.buildId)
					                        left join new_EmployeesInRoom r3 on (r3.roomId = r1.roomId)
					                        left join [dbo].new_RoomDetails r5 on (r1.roomId = r5.roomId) 
					                        left join cyclone.dbo.ROOM_2D r8 on (r1.roomFilename3ds = r8.Name)
				                                left join [dbo].[new_SubdivBuild] r9 on (r2.invNumber = r9.buildNumber)
					                            left join new_Subdivisions r7 on(r9.subdivId = r7.DepartmentId)
	                                    where (@buildNum is null or r2.invNumber = @buildNum) and (@roomId is null or r1.roomId=@roomId) and ";
                    string oper = "";
                    string action2 = declare + action;
                    //var predicate = PredicateBuilder.True<Room>();
                    //predicate = predicate.And(r => r.buildInvNumber.Contains(id) && r.subdivisionId == subdiv);//&& r.idChief == chief
                    string action3 = CreateQuery(action2, parsedData);
                    string query = SubstringQuery(action3);
                    using (AdmShipDataContext db3 = new AdmShipDataContext())
                    {
                        var rooms_pars = db3.ExecuteQuery<GetRoomByAllFields_oldResult>(query).OrderBy(o => o.roomId).ToList();//db3.GetTable<Room>().FromSql($"SELECT * FROM dbo.rooms({action})").ToList();
                        rooms = rooms_pars;
                    }
                }
                else
                {
                    using (AdmShipDataContext db3 = new AdmShipDataContext())
                    {
                        rooms = db3.GetRoomByAllFields_old(id, roomId).ToList();
                    }
                }
                foreach (var room in rooms)
                {
                    TableToFront2 tbl2 = new TableToFront2();
                    tbl2.OBJECTID = room.OBJECTID;
                    tbl2.roomFloor = room.roomFloor;
                    tbl2.roomNumber = room.uniqueRoomNumber;
                    tbl2.roomId = (int)room.roomId;
                    if (room.empFio != null)
                    {
                        tbl2.employeeFio = room.empFio;//.Split(',');
                    }
                    else
                    {
                        tbl2.employeeFio = /*new string[] {*/ string.Empty;
                    }
                    if (room.chiefFio != null)
                        tbl2.chiefFio = room.chiefFio;
                    else
                        tbl2.chiefFio = string.Empty;
                    tbl2.subdivisionId = room.DepartmentId;
                    tbl2.subdivision = room.SubdivisionName;
                    tbl2.depId = room.DepartmentId;
                    tbl2.buildInvNumber = room.invNumber;
                    tbl2.buildName = room.buildName;
                    tbl2.filename3ds = room.roomFilename3ds;
                    lst.Add(tbl2);
                }

                foreach (var t in lst)
                {
                    if (t.subdivisionId == null)
                        t.isSubdivNull = 0;
                    else
                        t.isSubdivNull = 1;
                    if (t.OBJECTID == null)
                        t.isVisible = 0;
                    else
                        t.isVisible = 1;
                }
                objToReturn.flag = false;
                objToReturn.list = lst;
                //dict.Add(objToReturn);
            }
        
        return Json(objToReturn);
        }

        private string CreateQuery(string action2, List<FromFilter> parsedData) {
            string oper = "";
            for (int i = 0, j = 1; i < parsedData.Count; i++, j++)
            {
                var item = parsedData[i];
                FromFilter item2 = new FromFilter();
                if (j <= parsedData.Count - 1)
                {
                    item2 = parsedData[j];
                }
                switch (item.parametr)
                {
                    case "floor":
                        switch (item.type)
                        {
                            case "contains":
                                action2 += "r5.roomFloor like '%" + item.value + "%'";
                                break;
                            case "eq":
                                action2 += "r5.roomFloor = '" + item.value + "'";
                                break;
                            case "gt":
                                action2 += "r5.roomFloor > '" + item.value + "'";
                                break;
                            case "lt":
                                action2 += "r5.roomFloor < '" + item.value + "'";
                                break;
                        }
                        break;
                    case "buildName":
                        switch (item.type)
                        {
                            case "contains":
                                action2 += "r2.buildName like '%" + item.value + "%'";
                                break;
                            case "eq":
                                action2 += "r2.buildName = '" + item.value + "'";
                                break;
                                //case "gt":
                                //    predicate = predicate.And(r => r.floor > Convert.ToDecimal(item.value));
                                //    break;
                                //case "lt":
                                //    predicate = predicate.And(r => r.floor < Convert.ToDecimal(item.value));
                                //    break;
                        }
                        break;
                    case "roomFunction":
                        switch (item.type)
                        {
                            case "contains":
                                action2 += "r5.roomFunction like '%" + item.value + "%'";
                                break;
                            case "eq":
                                action2 += "r5.roomFunction = '" + item.value + "'";
                                break;
                                //case "gt":
                                //    predicate = predicate.And(r => r.floor > Convert.ToDecimal(item.value));
                                //    break;
                                //case "lt":
                                //    predicate = predicate.And(r => r.floor < Convert.ToDecimal(item.value));
                                //    break;
                        }
                        break;
                    case "roomSquare":

                        switch (item.type)
                        {
                            //case "contains":
                            //    action2 += "room_square like '%" + item.value + "%' and ";
                            //    break;
                            case "eq":
                                action2 += "r5.roomSquare = " + item.value;
                                break;
                            case "gt":
                                action2 += "r5.roomSquare > " + item.value;
                                break;
                            case "lt":
                                action2 += "r5.roomSquare < " + item.value;
                                break;
                        }
                        break;
                    case "totalLiveSpace":
                        switch (item.type)
                        {
                            //case "contains":
                            //    action2 += "total_living_space like '%" + item.value + "%' and ";
                            //    break;
                            case "eq":
                                action2 += "r5.totalLivingSpace = " + item.value;
                                break;
                            case "gt":
                                action2 += "r5.totalLivingSpace > " + item.value;
                                break;
                            case "lt":
                                action2 += "r5.totalLivingSpace < " + item.value;
                                break;
                        }
                        break;
                    case "specialFunction":
                        switch (item.type)
                        {
                            //case "contains":
                            //    predicate = predicate.And(r => r.floor.Contains(item.value));
                            //    break;
                            case "eq":
                                action2 += "r5.roomSpecFunction = " + item.value;
                                break;
                            case "gt":
                                action2 += "r5.roomSpecFunction > " + item.value;
                                break;
                            case "lt":
                                action2 += "r5.roomSpecFunction < " + item.value;
                                break;
                        }
                        break;
                    case "ancillarySquare":
                        switch (item.type)
                        {
                            //case "contains":
                            //    predicate = predicate.And(r => r.floor.Contains(item.value));
                            //    break;
                            case "eq":
                                action2 += "r5.ancillarySquare = " + item.value;
                                break;
                            case "gt":
                                action2 += "r5.ancillarySquare > " + item.value;
                                break;
                            case "lt":
                                action2 += "r5.ancillarySquare < " + item.value;
                                break;
                        }
                        break;
                    case "roomHeight":
                        switch (item.type)
                        {
                            //case "contains":
                            //    predicate = predicate.And(r => r.floor.Contains(item.value));
                            //    break;
                            case "eq":
                                action2 += "r5.roomHeight = '" + item.value.ToString() + "'";
                                break;
                            case "gt":
                                action2 += "r5.roomHeight > '" + item.value.ToString() + "'";
                                break;
                            case "lt":
                                action2 += "r5.roomHeight < '" + item.value.ToString() + "'";
                                break;
                        }
                        break;
                    case "balconySquare":
                        switch (item.type)
                        {
                            //case "contains":
                            //    predicate = predicate.And(r => r.floor.Contains(item.value));
                            //    break;
                            case "eq":
                                action2 += "r5.balconySquare = '" + item.value + "'";
                                break;
                            case "gt":
                                action2 += "r5.balconySquare > '" + item.value + "'";
                                break;
                            case "lt":
                                action2 += "r5.balconySquare < '" + item.value + "'";
                                break;
                        }
                        break;
                }
                if (i != parsedData.Count - 1)
                {
                    if (item.parametr.Equals(item2.parametr))
                    {
                        oper = "or";
                    }
                    else
                    {
                        oper = "and";
                    }
                    action2 += " " + oper + " ";
                }
            }
            return action2;
        }

        [HttpPost]
        public JsonResult Index3D(int subdiv, string id = "", int? chief = null, int? roomId = null, string data = "", int? emplId = null)
        {
            //interfereDataBases();
            List<GetRoomByAllFieldsResult> rooms;
            if (data != "")
            {
                List<FromFilter> parsedData = (List<FromFilter>)JsonConvert.DeserializeObject(data, typeof(List<FromFilter>));
                string declare = "declare @subDivId int = " + subdiv + ";";
                if (id == "") {
                    declare+= "declare @buildNum nvarchar(6);"; 
                }
                else
                    declare+= "declare @buildNum nvarchar(6) = '" + id + "';";
                if (roomId == null) {
                    declare += "declare @roomId int;";
                }
                else {
                    declare += "declare @roomId int= " + roomId + ";";
                }
                if (chief == null)
                {
                    declare += "declare @chiefId int;";
                }
                else
                {
                    declare += "declare @chiefId int= " + chief + ";";
                }
                if (emplId == null)
                {
                    declare += "declare @empId int;";
                }
                else
                {
                    declare += "declare @empId int= " + emplId + ";";
                }
                string action = @"SELECT distinct r7.SubdivisionName, r7.DepartmentId, r2.invNumber, r2.buildName, r5.roomFloor, r1.roomId, r1.roomFilename3ds, r1.uniqueRoomNumber, r8.OBJECTID,
	                [dbo].GetChiefFio(r1.roomId, @subDivId) chiefFio, [dbo].GetEmpFio(r1.roomId, @subDivId) empFio from [dbo].new_Subdivisions r7
	                left join [dbo].[new_SubdivBuild] r9 on (r7.DepartmentId = r9.subdivId)
	                left join [dbo].new_Buildings r2 on(r9.buildNumber = r2.invNumber)
	                join [dbo].new_Rooms r1  on(r2.buildId = r1.buildId)
                    left join new_EmployeesInRoom r3 on (r3.roomId = r1.roomId)
	                left join [dbo].new_RoomDetails r5 on (r1.roomId = r5.roomId) 
	                left join cyclone.dbo.ROOM_3D r8 on (r1.roomFilename3ds = r8.Name)
	                where (@subDivId is null or r7.DepartmentId = @subDivId)" +
                    " and (@buildNum is null or r2.invNumber = @buildNum)" +
                    " and (@roomId is null or r1.roomId = @roomId)" +
                    " and (@chiefId is null or r3.chiefId = @chiefId) and (@empId is null or r3.empId = @empId) and ";
                string oper = "";
                string action2 = declare + action;
                //var predicate = PredicateBuilder.True<Room>();
                //predicate = predicate.And(r => r.buildInvNumber.Contains(id) && r.subdivisionId == subdiv);//&& r.idChief == chief
                for (int i = 0, j = 1; i < parsedData.Count; i++, j++)
                {
                    var item = parsedData[i];
                    FromFilter item2 = new FromFilter();
                    if (j <= parsedData.Count - 1)
                    {
                        item2 = parsedData[j];
                    }
                    switch (item.parametr)
                    {
                        case "floor":
                            switch (item.type)
                            {
                                case "contains":
                                    action2 += "r5.roomFloor like '%" + item.value + "%'";
                                    break;
                                case "eq":
                                    action2 += "r5.roomFloor = '" + item.value + "'";
                                    break;
                                case "gt":
                                    action2 += "r5.roomFloor > '" + item.value + "'";
                                    break;
                                case "lt":
                                    action2 += "r5.roomFloor < '" + item.value + "'";
                                    break;
                            }
                            break;
                        case "buildName":
                            switch (item.type)
                            {
                                case "contains":
                                    action2 += "r2.buildName like '%" + item.value + "%'";
                                    break;
                                case "eq":
                                    action2 += "r2.buildName = '" + item.value + "'";
                                    break;
                                    //case "gt":
                                    //    predicate = predicate.And(r => r.floor > Convert.ToDecimal(item.value));
                                    //    break;
                                    //case "lt":
                                    //    predicate = predicate.And(r => r.floor < Convert.ToDecimal(item.value));
                                    //    break;
                            }
                            break;
                        case "roomFunction":
                            switch (item.type)
                            {
                                case "contains":
                                    action2 += "r5.roomFunction like '%" + item.value + "%'";
                                    break;
                                case "eq":
                                    action2 += "r5.roomFunction = '" + item.value + "'";
                                    break;
                                    //case "gt":
                                    //    predicate = predicate.And(r => r.floor > Convert.ToDecimal(item.value));
                                    //    break;
                                    //case "lt":
                                    //    predicate = predicate.And(r => r.floor < Convert.ToDecimal(item.value));
                                    //    break;
                            }
                            break;
                        case "roomSquare":
                            decimal value = Convert.ToDecimal(item.value, CultureInfo.InvariantCulture);
                            switch (item.type)
                            {
                                //case "contains":
                                //    action2 += "room_square like '%" + item.value + "%' and ";
                                //    break;
                                case "eq":
                                    action2 += "r5.roomSquare = " + item.value;
                                    break;
                                case "gt":
                                    action2 += "r5.roomSquare > " + item.value;
                                    break;
                                case "lt":
                                    action2 += "r5.roomSquare < " + item.value;
                                    break;
                            }
                            break;
                        case "totalLiveSpace":
                            decimal value2 = Convert.ToDecimal(item.value, CultureInfo.InvariantCulture);
                            switch (item.type)
                            {
                                //case "contains":
                                //    action2 += "total_living_space like '%" + item.value + "%' and ";
                                //    break;
                                case "eq":
                                    action2 += "r5.totalLivingSpace = " + item.value;
                                    break;
                                case "gt":
                                    action2 += "r5.totalLivingSpace > " + item.value;
                                    break;
                                case "lt":
                                    action2 += "r5.totalLivingSpace < " + item.value;
                                    break;
                            }
                            break;
                        case "specialFunction":
                            decimal value3 = Convert.ToDecimal(item.value, CultureInfo.InvariantCulture);
                            switch (item.type)
                            {
                                //case "contains":
                                //    predicate = predicate.And(r => r.floor.Contains(item.value));
                                //    break;
                                case "eq":
                                    action2 += "r5.roomSpecFunction = " + item.value;
                                    break;
                                case "gt":
                                    action2 += "r5.roomSpecFunction > " + item.value;
                                    break;
                                case "lt":
                                    action2 += "r5.roomSpecFunction < " + item.value;
                                    break;
                            }
                            break;
                        case "ancillarySquare":
                            decimal value4 = Convert.ToDecimal(item.value, CultureInfo.InvariantCulture);
                            switch (item.type)
                            {
                                //case "contains":
                                //    predicate = predicate.And(r => r.floor.Contains(item.value));
                                //    break;
                                case "eq":
                                    action2 += "r5.ancillarySquare = " + item.value;
                                    break;
                                case "gt":
                                    action2 += "r5.ancillarySquare > " + item.value;
                                    break;
                                case "lt":
                                    action2 += "r5.ancillarySquare < " + item.value;
                                    break;
                            }
                            break;
                        case "roomHeight":
                            double value5 = Convert.ToDouble(item.value, CultureInfo.InvariantCulture);
                            switch (item.type)
                            {
                                //case "contains":
                                //    predicate = predicate.And(r => r.floor.Contains(item.value));
                                //    break;
                                case "eq":
                                    action2 += "r5.roomHeight = '" + item.value.ToString() + "'";
                                    break;
                                case "gt":
                                    action2 += "r5.roomHeight > '" + item.value.ToString() + "'";
                                    break;
                                case "lt":
                                    action2 += "r5.roomHeight < '" + item.value.ToString() + "'";
                                    break;
                            }
                            break;
                        case "balconySquare":
                            decimal value6 = Convert.ToDecimal(item.value, CultureInfo.InvariantCulture);
                            switch (item.type)
                            {
                                //case "contains":
                                //    predicate = predicate.And(r => r.floor.Contains(item.value));
                                //    break;
                                case "eq":
                                    action2 += "r5.balconySquare = '" + item.value + "'";
                                    break;
                                case "gt":
                                    action2 += "r5.balconySquare > '" + item.value + "'";
                                    break;
                                case "lt":
                                    action2 += "r5.balconySquare < '" + item.value + "'";
                                    break;
                            }
                            break;
                    }
                    if (i != parsedData.Count - 1)
                    {
                        if (item.parametr.Equals(item2.parametr))
                        {
                            oper = "or";
                        }
                        else
                        {
                            oper = "and";
                        }
                        action2 += " " + oper + " ";
                    }
                }
                string query = SubstringQuery(action2);
                using (AdmShipDataContext db3 = new AdmShipDataContext())
                {
                    var rooms_pars = db3.ExecuteQuery<GetRoomByAllFieldsResult>(query).OrderBy(o => o.roomId).ToList();//db3.GetTable<Room>().FromSql($"SELECT * FROM dbo.rooms({action})").ToList();
                    rooms = rooms_pars;
                }
            }
            else
            {
                using (AdmShipDataContext db3 = new AdmShipDataContext())
                {
                    if(id!="")
                        rooms = db3.GetRoomByAllFields(subdiv, id, roomId, chief, emplId).OrderBy(o => o.roomId).ToList();//GetTable<Room>().Where(r => r.subdivisionId == subdiv).OrderBy(o => o.uniqueRoomNumber).ToList();&& r.idChief == chief - добавим позже
                    else
                        rooms = db3.GetRoomByAllFields(subdiv, null, roomId, chief, emplId).OrderBy(o => o.roomId).ToList();
                }
            }
            List<TableToFront2> lst = new List<TableToFront2>();
            List<TableToFront2> lstChief = new List<TableToFront2>();
            Dictionary<string, List<TableToFront2>> dict = new Dictionary<string, List<TableToFront2>>();
            
            
            foreach (var room in rooms)
            {
                TableToFront2 tbl2 = new TableToFront2();
                tbl2.roomFloor = room.roomFloor;
                tbl2.roomNumber = room.uniqueRoomNumber;
                tbl2.roomId = room.roomId;
                tbl2.OBJECTID = room.OBJECTID;
                tbl2.subdivisionId = (int)room.DepartmentId;
                tbl2.subdivision = room.SubdivisionName;
                tbl2.depId = (int)room.DepartmentId;
                tbl2.buildInvNumber = room.invNumber;
                tbl2.buildName = room.buildName;
                tbl2.filename3ds = room.roomFilename3ds;
                if (room.chiefFio != null)
                    tbl2.chiefFio = room.chiefFio;
                else
                    tbl2.chiefFio = string.Empty;
                if (room.empFio != null)
                {
                    tbl2.employeeFio = room.empFio;//.Split(',');
                }
                else
                {
                    tbl2.employeeFio = /*new string[] {*/ string.Empty;
                }
                lst.Add(tbl2);
            }
            foreach (var t in lst) {
                if (t.OBJECTID != null)
                {
                    t.isVisible = 1;
                }
                else
                {
                    t.isVisible = 0;
                }
            }
            //int point3 = 0;
            dict.Add("list", lst);
            var jsonRes = Json(dict);
            jsonRes.MaxJsonLength = int.MaxValue;
            return Json(JsonConvert.SerializeObject(dict));
        }
        private string SubstringQuery(string query)
        {

            var andArgs = query.Split(new[] { "and", "AND", "And", "anD" }, StringSplitOptions.RemoveEmptyEntries);
            var processed = string.Join(" and ", andArgs.Select(c =>
            {
                var trim = c.Trim();
                if (!trim.ToLower().Contains("where"))
                {
                    if (!trim.ToLower().Contains("or"))
                        return trim;
                }
                else
                {
                    int index = trim.IndexOf("where ");
                    string newTrim = trim.Insert(index + "where ".Length, "(");
                    newTrim += ")";
                    trim = newTrim;
                }
                return !trim.StartsWith("(") && !trim.EndsWith(")") ? $"({trim})" : trim;
            }));
            return processed;
        }


        [HttpPost]
        public JsonResult UpdateGridNew(string statementType, int roomId, int id, int subdivId, int empId=0, string empText="", string empPhone = "", int chiefId=0, string chiefFio="") {
            JsonResult result = null;
            char[] trimChars = new char[] { '_', '-' };
            string phone = empPhone.TrimEnd(trimChars);
            using (AdmShipDataContext db3 = new AdmShipDataContext()) {
                db3.MasterInsertUpdateDelete(id, roomId, subdivId, empId, empText, chiefId, chiefFio, phone, statementType);
            }
            result = GetRoomEmployees(roomId, subdivId);
            return result;
        }
        public JsonResult GetRoomEmployees(int roomId, int subdiv)
        {
            JsonResult jsonResult = new JsonResult();
            using (AdmShipDataContext db3 = new AdmShipDataContext())
            {
                var rooms = db3.GetRoomsForEdit(roomId, subdiv).ToList();
                jsonResult = Json(rooms);
            }
            return jsonResult;
        }
        [HttpPost]
        public JsonResult IsEmplInEditor(int id)
        {
            JsonResult jsonResult = new JsonResult();
            using (AdmShipDataContext db3 = new AdmShipDataContext())
            {
                var emps = db3.IsEmplEdit(id).ToList();
                jsonResult = Json(emps);
            }
            return jsonResult;
        }
        [HttpPost]
        public JsonResult EmployeeChanged(int empId)
        {
            dynamic room;
            using (AdmShipDataContext db3 = new AdmShipDataContext())
            {
                room = db3.GetBuildsByEmpl(empId).ToList();
            }
            return Json(room);
        }
        [HttpPost]
        public JsonResult ChiefChanged(int id)
        {
            dynamic room;
            using (AdmShipDataContext db3 = new AdmShipDataContext())
            {
                room = db3.GetBuildsByChief(id).ToList();
            }
            return Json(room);
        }
        [HttpPost]
        public JsonResult SubdivChanged(int id)
        {
            string json = "";
            JsonResult jsonResult = new JsonResult();

            using (AdmShipDataContext db3 = new AdmShipDataContext())
            {
                var builds = db3.GetBuildings3d().Where(s => s.subdivId == id).ToList();
            
                if (builds != null)
                {
                    json = JsonConvert.SerializeObject(builds, Formatting.Indented);
                }
                jsonResult = Json(json);
                jsonResult.MaxJsonLength = int.MaxValue;
            }
            return jsonResult;
        }
        [HttpPost]
        public JsonResult GetFilteringCombobox(string value)//переписать
        {
            string json = "";

            switch (value)
            {
                case "buildName":
                    Dictionary<string, List<FilterCombobox>> dict = new Dictionary<string, List<FilterCombobox>>();
                    List<FilterCombobox> lst = new List<FilterCombobox>();
                    using (AdmShipDataContext db3 = new AdmShipDataContext())
                    {
                        var result = db3.GetBuildsForFilter().ToList();
                        foreach (var b in result)
                        {
                            FilterCombobox fcb = new FilterCombobox()
                            {
                                value = b.buildName,
                                text = b.buildName
                            };
                            lst.Add(fcb);
                        }
                        dict.Add("list", lst.OrderBy(o => o.value).ToList());
                    }
                    json = JsonConvert.SerializeObject(dict, Formatting.Indented);
                    break;
                case "roomFunction":
                    Dictionary<string, List<FilterCombobox>> dict2 = new Dictionary<string, List<FilterCombobox>>();
                    List<FilterCombobox> lst2 = new List<FilterCombobox>();
                    using (AdmShipDataContext db3 = new AdmShipDataContext())
                    {
                        var result2 = db3.GetRoomFuncsForFilter().ToList();

                        foreach (var r in result2)
                        {

                            FilterCombobox fcb = new FilterCombobox()
                            {
                                value = r.roomFunction,
                                text = r.roomFunction
                            };
                            lst2.Add(fcb);
                        }
                    }
                    dict2.Add("list", lst2.OrderBy(o => o.value).ToList());
                    json = JsonConvert.SerializeObject(dict2, Formatting.Indented);
                    break;
                case "floor":
                    
                    Dictionary<string, List<FilterCombobox>> dict3 = new Dictionary<string, List<FilterCombobox>>();
                    List<FilterCombobox> lst3 = new List<FilterCombobox>();
                    using (AdmShipDataContext db3 = new AdmShipDataContext())
                    {
                        var result3 = db3.GetFloorsForFilter().ToList();
                        foreach (var r in result3)
                        {
                            FilterCombobox fcb = new FilterCombobox()
                            {
                                value = r.roomFloor,
                                text = r.roomFloor
                            };
                            lst3.Add(fcb);
                        }
                    }
                    dict3.Add("list", lst3.OrderBy(o => o.value).ToList());
                    json = JsonConvert.SerializeObject(dict3, Formatting.Indented);
                    break;
            }
            return Json(json);
        }
        [HttpPost]
        public JsonResult GetSingleRoom(int id)//ПЕРЕПИСАТЬ!!!!
        {
            Room2 room = new Room2();
            using (AdmShipDataContext db3 = new AdmShipDataContext())
            {
                var sRoom = db3.GetSingleRoom(id);
                foreach (var r in sRoom)
                {
                    room.floor = r.RoomFloor;
                    room.roomNumber = r.roomNumber;
                    room.planRoomNumber = r.roomPlanNumber;
                    room.roomFunction = r.roomFunction;
                    room.roomSquare = (decimal)r.roomSquare;
                    room.totalLiveSpace = (decimal)r.totalLivingSpace;
                    room.specialFunction = (decimal)r.roomSpecFunction;
                    room.ancillarySquare = (decimal)r.ancillarySquare;
                    room.roomHeight = r.roomHeight;
                    room.balconySquare = (decimal)r.balconySquare;
                    room.roomFilename = r.roomFilename3ds;
                }
            }
            return Json(room);
        }
        public JsonResult GetSingleRoom2D(string fileName)//ПЕРЕПИСАТЬ!!!!
        {
            Room2 room = new Room2();
            using (AdmShipDataContext db3 = new AdmShipDataContext())
            {
                var sRoom = db3.GetSingleRoom2D(fileName);
                foreach (var r in sRoom)
                {
                    room.floor = r.RoomFloor;
                    room.roomNumber = r.roomNumber;
                    room.planRoomNumber = r.roomPlanNumber;
                    room.roomFunction = r.roomFunction;
                    room.roomSquare = (decimal)r.roomSquare;
                    room.totalLiveSpace = (decimal)r.totalLivingSpace;
                    room.specialFunction = (decimal)r.roomSpecFunction;
                    room.ancillarySquare = (decimal)r.ancillarySquare;
                    room.roomHeight = r.roomHeight;
                    room.balconySquare = (decimal)r.balconySquare;
                    room.roomFilename = r.roomFilename3ds;
                    room.roomId = r.roomId;
                }
            }
            return Json(room);
        }
        [HttpPost]
        public JsonResult ExcelExport(string[] data, int subdiv = 0)
        {
            //var numbers = data.Split(',');
            Dictionary<string, List<GetRoomsForExcelFullResult>> dict = new Dictionary<string, List<GetRoomsForExcelFullResult>>();
            List<GetRoomsForExcelFullResult> roomsToExport = new List<GetRoomsForExcelFullResult>();
            using (AdmShipDataContext db3 = new AdmShipDataContext())
            {
                var rooms = db3.GetRoomsForExcelFull().ToList();//db3.GetTable<Room>().Where(r => data.Contains(r.uniqueRoomNumber).ToList();

                if (data!=null)
                {
                    if (data.Length != 0)
                    {
                        if (subdiv != 0)
                            roomsToExport = rooms.Where(w => data.Any(a => w.uniqueRoomNumber == a) && w.DepartmentId == subdiv).ToList();
                        else
                            roomsToExport = rooms.Where(w => data.Any(a => w.uniqueRoomNumber == a)).ToList();
                    }
                }
            }
            dict.Add("list", roomsToExport);
            return Json(dict);
        }

        [HttpGet]
        public JsonResult CheckSubdiv(int subdiv) {
            ObjectToReturn objToReturn = new ObjectToReturn();
            using (AdmShipDataContext dc = new AdmShipDataContext()) {
                var check = dc.IsAnyConnection().ToList();
                if (check.Any(x => x.subdivId == subdiv))
                {
                    objToReturn.flag = true;
                }
                else
                    objToReturn.flag = false;
            }
            return Json(objToReturn, JsonRequestBehavior.AllowGet);
        }
        //[HttpPost]
        //public JsonResult GetUniqueRoomNumber(string fileName)
        //{
        //    var unique = db3.GetTable<Room>().Where(w => w.roomFilename == fileName).Select(s => s.uniqueRoomNumber);
        //    return Json(unique);
        //}
    }
}