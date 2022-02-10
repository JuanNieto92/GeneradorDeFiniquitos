using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GeneradorDeFiniquitos.Controllers
{
    public class FiniquitoController : Controller
    {
        // GET: Finiquito
        public ActionResult Index()
        {
            return View();
        }
    }
}