using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WillDo.Models;

namespace WillDo.Controllers
{
    public class PersonController : Controller
    {
        // GET: /MyCricketer/
        public ActionResult Index()
        {
            return View("PersonBegin");
        }
    }
}